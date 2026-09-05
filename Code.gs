function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var result = {};
  
  if (action === 'getProducts') {
    result = getProducts();
  } else if (action === 'getAdminStats') {
    result = getAdminStats();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var postData = JSON.parse(e.postData.contents);
  var orderId = submitOrder(postData.nama, postData.wa, postData.paket, postData.total);
  
  return ContentService.createTextOutput(JSON.stringify({ orderId: orderId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Produk');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  data.shift();
  return data;
}

function submitOrder(nama, wa, paket, total) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pesanan');
  var orderId = 'PYNI-' + Math.floor(100000 + Math.random() * 900000);
  var tgl = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  
  sheet.appendRow([orderId, tgl, nama, wa, paket, total, 'Menunggu Pembayaran']);
  return orderId;
}

function getAdminStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prodSheet = ss.getSheetByName('Produk');
  var orderSheet = ss.getSheetByName('Pesanan');
  
  var totalProduk = prodSheet ? Math.max(0, prodSheet.getLastRow() - 1) : 0;
  var orders = orderSheet ? orderSheet.getDataRange().getValues() : [];
  if (orders.length > 0) orders.shift();
  
  var totalOmzet = 0;
  orders.forEach(function(row) {
    totalOmzet += Number(row[5]) || 0;
  });

  return {
    totalProduk: totalProduk,
    totalPesanan: orders.length,
    totalOmzet: totalOmzet,
    riwayat: orders.slice(-10).reverse()
  };
}
