/**
 * ========================================
 * HELPER FUNCTIONS
 * ========================================
 */

/**
 * Generates a unique Order ID based on timestamp
 * Format: ORD-YYYYMMDD-###
 */
function generateOrderID(timestamp) {
  // Convert timestamp to date
  var date = new Date(timestamp);
  
  // Format: YYYYMMDD
  var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
  
  // Get the Orders sheet to find the last order number for today
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ordersSheet = ss.getSheetByName("Orders");
  var lastRow = ordersSheet.getLastRow();
  
  // Count UNIQUE order IDs for today (not total rows!)
  var todayOrders = 0;
  var todayPrefix = "ORD-" + dateStr;
  var seenOrderIDs = {};  // Track unique IDs
  
  if (lastRow > 1) {
    var orderIDs = ordersSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (var i = 0; i < orderIDs.length; i++) {
      var orderID = orderIDs[i][0].toString();
      
      // Check if it's from today AND we haven't counted it yet
      if (orderID.startsWith(todayPrefix) && !seenOrderIDs[orderID]) {
        seenOrderIDs[orderID] = true;  // Mark as seen
        todayOrders++;
      }
    }
  }
  
  // Increment counter
  var orderNumber = (todayOrders + 1).toString().padStart(3, '0');
  
  // Return formatted Order ID
  return "ORD-" + dateStr + "-" + orderNumber;
}

/**
 * Looks up delivery cost based on location
 */
function getDeliveryCost(deliveryLocation, homeAddress) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deliverySheet = ss.getSheetByName("Delivery_Zones");
  
  // Check if it's a home delivery
  if (deliveryLocation === "HOME DELIVERY/OTHER (FILL BELOW)") {
    Logger.log("Home delivery to: " + homeAddress + " - Default P50");
    return 50;
  }
  
  // For standard locations, look up in Delivery_Zones
  var data = deliverySheet.getDataRange().getValues();
  
  // Skip header row and search for matching location
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === deliveryLocation) {
      return data[i][1];
    }
  }
  
  // If location not found, default to 50
  Logger.log("Location not found: " + deliveryLocation + " - Defaulting to P50");
  return 50;
}

/**
 * Looks up product price from Product_Master sheet
 */
function getProductPrice(formHeader) {
  // Safety check
  if (!formHeader) {
    Logger.log("ERROR: formHeader is undefined or empty!");
    return 0;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var productSheet = ss.getSheetByName("Product_Master");
  
  // Strip the price from form header
  var productName = formHeader.split("- P")[0].trim();
  
  // Get all data from Product_Master
  var data = productSheet.getDataRange().getValues();
  
  // Search for matching product (skip header row)
  for (var i = 1; i < data.length; i++) {
    var masterProductName = data[i][0];  // Column A
    
    if (masterProductName === productName) {
      var sellingPrice = data[i][4];  // Column E
      return sellingPrice;
    }
  }
  
  // If product not found
  Logger.log("WARNING: Product not found: " + productName);
  return 0;
}

/**
 * ========================================
 * MAIN PROCESSING FUNCTIONS
 * ========================================
 */

/**
 * Process a specific row from Form_Responses
 */
function processSpecificRow(rowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("Form_Responses");
  var ordersSheet = ss.getSheetByName("Orders");
  
  var lastColumn = formSheet.getLastColumn();
  var formData = formSheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  var headers = formSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  // PART 1: Extract customer information
  var timestamp = formData[0];
  var firstName = formData[1];
  var surname = formData[2];
  var phone = formData[3];
  var deliveryLocation = formData[4];
  var homeAddress = formData[5];
  
  // PART 2: Find product columns and payment method dynamically
  var productStartColumn = 6;
  var productEndColumn = -1;
  var paymentMethodColumn = -1;
  
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === "PREFERRED PAYMENT METHOD") {
      productEndColumn = i - 1;
      paymentMethodColumn = i;
      break;
    }
  }
  
  if (productEndColumn === -1) {
    Logger.log("ERROR: Could not find PREFERRED PAYMENT METHOD column");
    return;
  }
  
  var paymentMethod = formData[paymentMethodColumn];
  
  // PART 3: Generate Order ID and get delivery cost
  var orderID = generateOrderID(timestamp);
  var deliveryCost = getDeliveryCost(deliveryLocation, homeAddress);
  
  Logger.log("Order ID: " + orderID + " | Customer: " + firstName + " " + surname);
  
  // PART 4: Loop through product columns
  var orderRows = [];
  var productsTotal = 0;
  
  for (var i = productStartColumn; i <= productEndColumn; i++) {
    var cellValue = formData[i];
    
    if (cellValue === "" || cellValue === null || cellValue === undefined) {
      continue;
    }
    
    var productHeader = headers[i];
    var productPrice = getProductPrice(productHeader);
    var productName = productHeader.split("- P")[0].trim();
    
    // Check if it's a number (simple quantity) or variant(s)
    if (!isNaN(cellValue) && cellValue !== "") {
      // It's a number - simple quantity
      var quantity = parseInt(cellValue);
      var lineTotal = productPrice * quantity;
      productsTotal += lineTotal;
      
      var orderRow = [
        orderID, timestamp, firstName, surname, phone,
        deliveryLocation, deliveryCost, productName, "",
        quantity, productPrice, lineTotal, paymentMethod, 0
      ];
      orderRows.push(orderRow);
      Logger.log("  - " + productName + " x" + quantity);
      
    } else {
      // It's a variant or multiple variants (comma-separated)
      var variants = String(cellValue).split(",");
      
      for (var v = 0; v < variants.length; v++) {
        var variant = variants[v].trim();
        
        var lineTotal = productPrice * 1;
        productsTotal += lineTotal;
        
        var orderRow = [
          orderID, timestamp, firstName, surname, phone,
          deliveryLocation, deliveryCost, productName, variant,
          1, productPrice, lineTotal, paymentMethod, 0
        ];
        orderRows.push(orderRow);
        Logger.log("  - " + productName + " (" + variant + ") x1");
      }
    }
  }
  
  // PART 5: Add delivery fee
  if (deliveryCost > 0) {
    var deliveryRow = [
      orderID, timestamp, firstName, surname, phone,
      deliveryLocation, deliveryCost, "DELIVERY FEE", "",
      1, deliveryCost, deliveryCost, paymentMethod, 0
    ];
    orderRows.push(deliveryRow);
  }
  
  // PART 6: Calculate grand total
  var grandTotal = productsTotal + deliveryCost;
  for (var j = 0; j < orderRows.length; j++) {
    orderRows[j][13] = grandTotal;
  }
  
  Logger.log("  TOTAL: P" + grandTotal);
  
  // PART 7: Write to Orders sheet
  if (orderRows.length > 0) {
    ordersSheet.getRange(ordersSheet.getLastRow() + 1, 1, orderRows.length, orderRows[0].length).setValues(orderRows);
  }
}

/**
 * Process ALL form responses
 * THIS IS THE FUNCTION YOU RUN!
 */
function processAllFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("Form_Responses");
  
  var lastRow = formSheet.getLastRow();
  
  if (lastRow < 2) {
    Logger.log("No form responses to process!");
    return;
  }
  
  Logger.log("=== Processing " + (lastRow - 1) + " form responses ===");
  
  for (var rowNum = 2; rowNum <= lastRow; rowNum++) {
    Logger.log("\n--- Processing row " + rowNum + " ---");
    processSpecificRow(rowNum);
  }
  
  Logger.log("\n=== ALL ORDERS PROCESSED ===");
}

/**
 * ========================================
 * TEST FUNCTIONS
 * ========================================
 */

function testOrderID() {
  var testTimestamp = "1/31/2026 20:05:26";
  var orderID = generateOrderID(testTimestamp);
  Logger.log("Generated Order ID: " + orderID);
}

function testDeliveryCost() {
  var cost1 = getDeliveryCost("ACACIA MALL- FREE", "");
  var cost2 = getDeliveryCost("AIRPORT JUNCTION- P30", "");
  var cost3 = getDeliveryCost("HOME DELIVERY/OTHER (FILL BELOW)", "Gaborone West");
  
  Logger.log("Acacia Mall: P" + cost1);
  Logger.log("Airport Junction: P" + cost2);
  Logger.log("Home Delivery: P" + cost3);
}

function testProductPrice() {
  var price1 = getProductPrice("TOAST KEYCHAIN- P35");
  var price2 = getProductPrice("BUTTONS PINS- P25");
  
  Logger.log("Toast Keychain: P" + price1);
  Logger.log("Button Pins: P" + price2);
}

/**
 * ========================================
 * AUTOMATIC TRIGGER FUNCTION
 * ========================================
 */

/**
 * This function runs automatically when a form is submitted
 * @param {object} e - The form submission event
 */
function onFormSubmit(e) {
  Logger.log("=== AUTOMATIC FORM SUBMISSION DETECTED ===");
  
  try {
    // Get the form sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("Form_Responses");
    
    // Get the last row (the new submission)
    var lastRow = formSheet.getLastRow();
    
    Logger.log("Processing new submission in row " + lastRow);
    
    // Process just this one row
    processSpecificRow(lastRow);
    
    Logger.log("✓ Order processed successfully!");
    
  } catch (error) {
    Logger.log("ERROR: " + error.toString());
    // Even if there's an error, don't crash - log it
  }
}

/**
 * ========================================
 * PACKING VIEW FUNCTIONS
 * ========================================
 */

/**
 * Refreshes the Packing View sheet with current orders
 * Groups products by order and shows status
 */
function refreshPackingView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ordersSheet = ss.getSheetByName("Orders");
  var packingSheet = ss.getSheetByName("Packing_View");
  
  // Create Packing_View sheet if it doesn't exist
  if (!packingSheet) {
    packingSheet = ss.insertSheet("Packing_View");
  }
  
  // Clear existing content
  packingSheet.clear();
  
  // Set up headers
  var headers = [
    "Completed", "Order_ID", "Order_Date", "Customer", "Phone", 
    "Delivery_Location", "Products_Summary", "Order_Total", "Payment_Method"
  ];
  packingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  packingSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  // Get all orders data
  var ordersData = ordersSheet.getDataRange().getValues();
  
  if (ordersData.length < 2) {
    Logger.log("No orders to display");
    return;
  }
  
  // Group orders by Order_ID
  var orderGroups = {};
  
  for (var i = 1; i < ordersData.length; i++) {
    var orderID = ordersData[i][0];
    var timestamp = ordersData[i][1];
    var firstName = ordersData[i][2];
    var surname = ordersData[i][3];
    var phone = ordersData[i][4];
    var deliveryLocation = ordersData[i][5];
    var productName = ordersData[i][7];
    var variant = ordersData[i][8];
    var quantity = ordersData[i][9];
    var orderTotal = ordersData[i][13];
    var paymentMethod = ordersData[i][12];
    
    // Initialize order group if first time seeing this order
    if (!orderGroups[orderID]) {
      orderGroups[orderID] = {
        orderID: orderID,
        timestamp: timestamp,
        customer: firstName + " " + surname,
        phone: phone,
        deliveryLocation: deliveryLocation,
        products: [],
        orderTotal: orderTotal,
        paymentMethod: paymentMethod
      };
    }
    
    // Add product to this order's list
    var productLine = productName;
    if (variant) {
      productLine += " (" + variant + ")";
    }
    productLine += " x" + quantity;
    
    orderGroups[orderID].products.push(productLine);
  }
  
  // Convert to array and write to sheet
  var packingData = [];
  
  for (var orderID in orderGroups) {
    var order = orderGroups[orderID];
    
    // Format date
    var date = new Date(order.timestamp);
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM dd, yyyy");
    
    // Combine products into one string
    var productsSummary = order.products.join("\n");
    
    var row = [
      false,  // Completed checkbox (unchecked by default)
      order.orderID,
      dateStr,
      order.customer,
      order.phone,
      order.deliveryLocation,
      productsSummary,
      order.orderTotal,
      order.paymentMethod
    ];
    
    packingData.push(row);
  }
  
  // Write data to sheet
  if (packingData.length > 0) {
    var startRow = 2;
    packingSheet.getRange(startRow, 1, packingData.length, packingData[0].length).setValues(packingData);
    
    // Format the Completed column as checkboxes
    packingSheet.getRange(startRow, 1, packingData.length, 1).insertCheckboxes();
    
    // Set text wrapping for products column
    packingSheet.getRange(startRow, 7, packingData.length, 1).setWrap(true);
    
    // Auto-resize columns
    packingSheet.autoResizeColumns(1, headers.length);
    
    // Set row heights
    for (var i = 0; i < packingData.length; i++) {
      var numProducts = packingData[i][6].split("\n").length;
      var rowHeight = Math.max(21, numProducts * 21);
      packingSheet.setRowHeight(startRow + i, rowHeight);
    }
  }
  
  Logger.log("✓ Packing View refreshed with " + packingData.length + " orders");
}

/**
 * Hides completed orders from Packing View
 * Run this after checking off orders
 */
function hideCompletedOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var packingSheet = ss.getSheetByName("Packing_View");
  
  if (!packingSheet) {
    Logger.log("Packing_View sheet not found");
    return;
  }
  
  var lastRow = packingSheet.getLastRow();
  
  if (lastRow < 2) {
    Logger.log("No orders to process");
    return;
  }
  
  // Get all data
  var data = packingSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  var hiddenCount = 0;
  
  // Loop through rows from bottom to top (so deleting doesn't mess up indices)
  for (var i = data.length - 1; i >= 0; i--) {
    var completed = data[i][0];  // Column A = Completed checkbox
    
    if (completed === true) {
      // Hide this row
      packingSheet.hideRows(i + 2);  // +2 because array is 0-indexed and row 1 is header
      hiddenCount++;
    }
  }
  
  Logger.log("✓ Hidden " + hiddenCount + " completed orders");
}

/**
 * Shows all hidden orders again
 */
function showAllOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var packingSheet = ss.getSheetByName("Packing_View");
  
  if (!packingSheet) {
    Logger.log("Packing_View sheet not found");
    return;
  }
  
  var lastRow = packingSheet.getLastRow();
  
  if (lastRow < 2) {
    return;
  }
  
  // Unhide all rows
  packingSheet.showRows(2, lastRow - 1);
  
  Logger.log("✓ All orders now visible");
}