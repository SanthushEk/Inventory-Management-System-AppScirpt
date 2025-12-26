function renderPurchesesPage() {
  return HtmlService.createHtmlOutputFromFile("purchases")
    .setTitle("Purchase Products")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Get all orders
function getOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1) // remove header
    .filter(r => r.some(c => c !== "" && c !== undefined))
    .map(row => ({
      orderId: row[0] || "",
      date: row[1] || "",
      supplierId: row[2] || "",
      supplierName: row[3] || "",
      model: row[4] || "",
      quantity: row[5] || ""
    }));
}

// Save new order
function saveOrder(order) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  if (!sheet) throw new Error("Orders sheet not found");

  const orderId = 'ORD-' + new Date().getTime();
  sheet.appendRow([
    orderId,
    order.date || "",
    order.supplierId || "",
    order.supplierName || "",
    order.model || "",
    order.quantity || ""
  ]);

  return { orderId, ...order };
}

// Update order by Order ID
function updateOrder(order) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === order.orderId) {
      sheet.getRange(i + 1, 2, 1, 5).setValues([[
        order.date || "",
        order.supplierId || "",
        order.supplierName || "",
        order.model || "",
        order.quantity || ""
      ]]);
      return order;
    }
  }
  throw new Error("Order not found");
}

// Delete order by Order ID
function deleteOrder(orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error("Order not found");
}
