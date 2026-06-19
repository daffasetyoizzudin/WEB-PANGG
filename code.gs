/**
 * PANGG.CASUALS - Premium Casual Wear Engine
 * Backend Google Apps Script (REST API, Sheets, and Drive Storage Integration)
 * VERSI PRODUKSI TEROPTIMALISASI - SEQUENTIAL MULTI-PHOTO UPLOAD CONTROL
 */

function doGet(e) {
  var callback = e.parameter.callback || e.parameter.Callback;
  try {
    var action = e.parameter.action || e.parameter.Action;
    
    if (!action) {
      return HtmlService.createTemplateFromFile("admin")
        .evaluate()
        .setTitle("PANGG.CASUALS - Admin Control Panel")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    initDatabaseAndFolders();
    
    if (action === "getTrackingData") {
      var code = e.parameter.code || e.parameter.Code;
      if (!code) {
        return jQueryResponse({ success: false, message: "Tracking Code tidak boleh kosong." }, callback);
      }
      return jQueryResponse(getTrackingDataByCode(ss, code), callback);
    }
    
    if (action === "getStoreData") {
      var settings = getSettings(ss);
      var products = getProducts(ss, false);
      return jQueryResponse({ success: true, settings: settings, products: products }, callback);
    } else if (action === "getProducts") {
      return jQueryResponse({ success: true, data: getProducts(ss, false) }, callback);
    } else if (action === "getProductsAdmin") {
      return jQueryResponse({ success: true, data: getProducts(ss, true) }, callback);
    } else if (action === "getSettings") {
      return jQueryResponse({ success: true, data: getSettings(ss) }, callback);
    } else if (action === "getOrders") {
      return jQueryResponse({ success: true, data: getOrders(ss) }, callback);
    } else if (action === "getDashboardStats") {
      return jQueryResponse({ success: true, data: getDashboardStats(ss) }, callback);
    }
    
    return jQueryResponse({ success: false, message: "Aksi GET '" + action + "' tidak dikenal." }, callback);
  } catch (err) {
    return jQueryResponse({ success: false, error: err.toString(), location: "doGet" }, callback);
  }
}

function doPost(e) {
  try {
    var contents = e.postData.contents;
    var payload = JSON.parse(contents);
    var action = payload.action;
    var data = payload.data;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    initDatabaseAndFolders();
    
    var result;
    if (action === "createOrderWithProof") {
      result = createOrderWithProof(ss, data);
    } else if (action === "uploadProof") {
      result = uploadProof(ss, payload);
    } else if (action === "updateOrderStatus") {
      result = updateOrderStatus(ss, data.OrderID, data.Status, data.Kurir, data.Resi);
    } else if (action === "createProduct") {
      result = createProduct(ss, data);
    } else if (action === "updateProduct") {
      result = updateProduct(ss, data);
    } else if (action === "deleteProduct") {
      result = deleteProduct(ss, data.ID);
    } else if (action === "updateSettings") {
      result = updateSettings(ss, data);
    } else {
      result = { success: false, message: "Aksi POST '" + action + "' tidak dikenal." };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString(), location: "doPost" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function jQueryResponse(obj, callback) {
  var sanitized = sanitizeForSerialization(obj);
  var jsonString = JSON.stringify(sanitized);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + jsonString + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

function initDatabaseAndFolders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetProduk = ss.getSheetByName("Produk");
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet("Produk");
    sheetProduk.appendRow([
      "ID", "Nama", "Harga", "Diskon", "HargaFinal", "Kategori", 
      "Ukuran", "Deskripsi", "Kondisi", "Stok", "Foto1", "Foto2", "Status",
      "Foto3", "Foto4", "Foto5", "Foto6", "Foto7", "Foto8", "Foto9", "Foto10"
    ]);
  }
  
  var sheetTransaksi = ss.getSheetByName("Transaksi");
  if (!sheetTransaksi) {
    sheetTransaksi = ss.insertSheet("Transaksi");
    sheetTransaksi.appendRow([
      "Tanggal", "OrderID", "NamaPembeli", "WhatsApp", "Alamat", 
      "Provinsi", "Catatan", "Produk", "Total", "BuktiTransfer", "Status", "TrackingCode", "Kurir", "Resi"
    ]);
  }
  
  var sheetConfig = ss.getSheetByName("Pengaturan");
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet("Pengaturan");
    sheetConfig.appendRow(["Kunci", "Nilai"]);
    sheetConfig.appendRow(["Logo", ""]);
    sheetConfig.appendRow(["Banner", "https://images.unsplash.com/photo-1707306354460-a292850fbf45?auto=format&fit=crop&q=80&w=1200"]);
    sheetConfig.appendRow(["QRIS", "https://images.unsplash.com/photo-1595079676339-1534801ad6cf?auto=format&fit=crop&q=80&w=200"]);
    sheetConfig.appendRow(["AdminWhatsapp", "6281234567890"]);
    sheetConfig.appendRow(["Instagram", "pangg.casuals"]);
    sheetConfig.appendRow(["TikTok", "pangg.casuals"]);
    sheetConfig.appendRow(["AdminUser", "admin"]);
    sheetConfig.appendRow(["AdminPass", "12345"]);
    sheetConfig.appendRow(["Vouchers", ""]);
  }
}

function generateRandomTrackingCode() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var code = "TRK";
  for (var i = 0; i < 6; i++) {
    var r = Math.floor(Math.random() * chars.length);
    code += chars.charAt(r);
  }
  return code;
}

function getTargetFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var newFolder = DriveApp.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

function createOrderWithProof(ss, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    
    var sheet = ss.getSheetByName("Transaksi");
    var orderId = data.OrderID || ("PC-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000));
    var trackingCode = generateRandomTrackingCode();
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    var fileUrl = "";
    if (data.fileData && data.fileData.base64) {
      fileUrl = saveImageToDriveDirect(data.fileData.base64, "Bukti_" + orderId, "Bukti Transfer - PANGG.CASUALS");
    }
    
    sheet.appendRow([
      timestamp,
      orderId,
      data.NamaPembeli,
      "'" + data.WhatsApp, 
      data.Alamat,
      data.Provinsi,
      data.Catatan,
      data.Produk,
      data.Total,
      fileUrl, 
      "MENUNGGU VERIFIKASI",
      trackingCode,
      "",
      ""
    ]);
    
    try {
      var productSheet = ss.getSheetByName("Produk");
      if (productSheet) {
        var prodRows = productSheet.getDataRange().getValues();
        var items = data.CartItems || [];
        items.forEach(function(item) {
          for (var i = 1; i < prodRows.length; i++) {
            if (prodRows[i][0].toString() === item.productId.toString()) {
              var currentStock = parseInt(prodRows[i][9]) || 0; 
              var newStock = Math.max(0, currentStock - (item.qty || 1));
              productSheet.getRange(i + 1, 10).setValue(newStock);
              break;
            }
          }
        });
      }
    } catch (stokErr) {
      Logger.log("Gagal memotong stok: " + stokErr.toString());
    }
    
    return { success: true, orderId: orderId, trackingCode: trackingCode, fileUrl: fileUrl };
  } finally {
    lock.releaseLock();
  }
}

function uploadProof(ss, postData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    
    var sheet = ss.getSheetByName("Transaksi");
    var orderId = postData.orderId;
    var fileData = postData.fileData;
    
    if (!fileData || !fileData.base64) {
      return { success: false, message: "Berkas gambar transfer kosong." };
    }
    
    var fileUrl = saveImageToDriveDirect(fileData.base64, "Bukti_" + orderId, "Bukti Transfer - PANGG.CASUALS");
    
    var data = sheet.getDataRange().getValues();
    var rowUpdated = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString() === orderId.toString()) {
        sheet.getRange(i + 1, 10).setValue(fileUrl); 
        sheet.getRange(i + 1, 11).setValue("MENUNGGU VERIFIKASI"); 
        rowUpdated = true;
        break;
      }
    }
    
    if (rowUpdated) {
      return { success: true, fileUrl: fileUrl };
    } else {
      return { success: false, message: "ID Pesanan " + orderId + " tidak ditemukan di database." };
    }
  } finally {
    lock.releaseLock();
  }
}

function getSettings(ss) {
  var sheet = ss.getSheetByName("Pengaturan");
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key) {
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
      } else if (val === null || val === undefined) {
        val = "";
      } else {
        val = val.toString();
      }
      settings[key.toString().trim()] = val;
    }
  }
  return settings;
}

function getProducts(ss, includeArchived) {
  var sheet = ss.getSheetByName("Produk");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var products = [];
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var item = {};
    var hasId = false;
    
    for (var c = 0; c < headers.length; c++) {
      var headerName = headers[c];
      if (headerName) {
        var val = row[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
        } else if (val === null || val === undefined) {
          val = "";
        } else {
          val = val.toString();
        }
        item[headerName] = val;
        if (headerName === "ID" && val) {
          hasId = true;
        }
      }
    }
    
    if (hasId) {
      if (includeArchived || (item.Status !== "ARSIP" && item.Status !== "ARCHIVED")) {
        products.push(item);
      }
    }
  }
  return products;
}

function getOrders(ss) {
  var sheet = ss.getSheetByName("Transaksi");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var orders = [];
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var item = {};
    var hasId = false;
    
    for (var c = 0; c < headers.length; c++) {
      var headerName = headers[c];
      if (headerName) {
        var val = row[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
        } else if (val === null || val === undefined) {
          val = "";
        } else {
          val = val.toString();
        }
        item[headerName] = val;
        if (headerName === "OrderID" && val) {
          hasId = true;
        }
      }
    }
    if (hasId) {
      orders.push(item);
    }
  }
  return orders;
}

function getDashboardStats(ss) {
  var sheetProduk = ss.getSheetByName("Produk");
  var sheetTransaksi = ss.getSheetByName("Transaksi");
  
  var totalProducts = 0;
  var soldOutProducts = 0;
  if (sheetProduk) {
    var prodData = sheetProduk.getDataRange().getValues();
    for (var i = 1; i < prodData.length; i++) {
      if (prodData[i][0]) { 
        totalProducts++;
        var stok = parseInt(prodData[i][9]) || 0; 
        if (stok <= 0) {
          soldOutProducts++;
        }
      }
    }
  }
  
  var totalOrders = 0;
  var pendingOrders = 0;
  if (sheetTransaksi) {
    var orderData = sheetTransaksi.getDataRange().getValues();
    for (var j = 1; j < orderData.length; j++) {
      if (orderData[j][1]) { 
        totalOrders++;
        var status = (orderData[j][10] || "").toString().toUpperCase(); 
        if (status === "PENDING" || status === "MENUNGGU_PEMBAYARAN" || status === "MENUNGGU VERIFIKASI") {
          pendingOrders++;
        }
      }
    }
  }
  
  return {
    totalProducts: totalProducts,
    totalOrders: totalOrders,
    pendingOrders: pendingOrders,
    soldOutProducts: soldOutProducts
  };
}

// PRODUK BARU TANPA BERKAS FOTO DAHULU (DIUNGGAH BERTURUT-TURUT DI FUNGSI BERBEDA)
function createProduct(ss, data) {
  var sheet = ss.getSheetByName("Produk");
  var id = "PROD-" + Math.floor(100000 + Math.random() * 900000);
  var finalPrice = parseFloat(data.Harga) * (1 - (parseFloat(data.Diskon) || 0) / 100);

  sheet.appendRow([
    id,
    data.Nama,
    data.Harga,
    data.Diskon || 0,
    finalPrice,
    data.Kategori,
    data.Ukuran,
    data.Deskripsi,
    data.Kondisi,
    data.Stok,
    "", // Foto1
    "", // Foto2
    data.Status || "AKTIF",
    "", // Foto3
    "", // Foto4
    "", // Foto5
    "", // Foto6
    "", // Foto7
    "", // Foto8
    "", // Foto9
    ""  // Foto10
  ]);
  return { success: true, id: id, message: "Rincian produk berhasil ditambahkan dengan ID: " + id };
}

// UPDATE PRODUK TANPA MENIMPA FOTO YANG LAMA (FOTO DIUPDATE SPESIFIK JIKA ADA)
function updateProduct(ss, data) {
  var sheet = ss.getSheetByName("Produk");
  var rows = sheet.getDataRange().getValues();
  var id = data.ID;
  var rowIndex = -1;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: "Produk tidak ditemukan." };
  }
  
  var finalPrice = parseFloat(data.Harga) * (1 - (parseFloat(data.Diskon) || 0) / 100);
  
  sheet.getRange(rowIndex, 2).setValue(data.Nama);
  sheet.getRange(rowIndex, 3).setValue(data.Harga);
  sheet.getRange(rowIndex, 4).setValue(data.Diskon || 0);
  sheet.getRange(rowIndex, 5).setValue(finalPrice);
  sheet.getRange(rowIndex, 6).setValue(data.Kategori);
  sheet.getRange(rowIndex, 7).setValue(data.Ukuran);
  sheet.getRange(rowIndex, 8).setValue(data.Deskripsi);
  sheet.getRange(rowIndex, 9).setValue(data.Kondisi);
  sheet.getRange(rowIndex, 10).setValue(data.Stok);
  sheet.getRange(rowIndex, 13).setValue(data.Status || "AKTIF");
  
  return { success: true, id: id, message: "Rincian produk berhasil diperbarui." };
}

// INTEGRASI UNGGAH FOTO SECARA SEQUENTIAL (SATU PER SATU) UNTUK MENCEGAH TIMEOUT
function uploadProductPhotoGAS(productId, photoIndex, base64DataStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Produk");
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === productId.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Produk tidak ditemukan saat upload foto." };
    }

    if (!base64DataStr) {
      return { success: true, message: "Foto kosong diabaikan." };
    }

    // Jika berupa URL lama (bukan base64), biarkan tidak usah diupload ulang
    if (base64DataStr.indexOf(";base64,") === -1 && base64DataStr.startsWith("http")) {
      return { success: true, message: "Foto berupa URL dipertahankan." };
    }

    var folder = getTargetFolder("Foto Produk - PANGG.CASUALS");
    var fileUrl = saveImageToDriveWithFolder(base64DataStr, "Foto" + photoIndex, productId, folder);

    // Pemetaan indeks kolom Foto1 s.d. Foto10 di Google Sheets
    var colIndex = 11; // Default Foto1
    if (photoIndex === 1) colIndex = 11;
    else if (photoIndex === 2) colIndex = 12;
    else if (photoIndex >= 3 && photoIndex <= 10) {
      colIndex = 11 + photoIndex; // Foto3 -> 14, Foto4 -> 15, dst.
    }

    sheet.getRange(rowIndex, colIndex).setValue(fileUrl);
    return { success: true, fileUrl: fileUrl, message: "Foto " + photoIndex + " berhasil diunggah." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function saveImageToDriveDirect(base64DataStr, prefix, folderName) {
  var folder = getTargetFolder(folderName);
  return saveImageToDriveWithFolder(base64DataStr, prefix, "ONE_SHOT", folder);
}

function saveImageToDriveWithFolder(base64DataStr, prefix, productId, folder) {
  if (!base64DataStr) return "";
  var str = base64DataStr.toString().trim();
  
  if (!str.startsWith("data:") || str.indexOf(";base64,") === -1) {
    return str; 
  }
  try {
    var base64Parts = str.split(",");
    var contentType = "image/jpeg";
    var base64Data = base64Parts[0];
    
    if (base64Parts.length > 1) {
      contentType = base64Parts[0].split(";")[0].split(":")[1];
      base64Data = base64Parts[1];
    }
    
    var decodedFile = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decodedFile, contentType, prefix + "_" + productId + "_" + new Date().getTime() + ".jpg");
    
    var file = folder.createFile(blob);
    return "https://lh3.googleusercontent.com/u/0/d/" + file.getId();
  } catch(e) {
    Logger.log("Gagal memproses unggah foto ke Drive: " + e.toString());
    return "";
  }
}

function deleteProduct(ss, id) {
  var sheet = ss.getSheetByName("Produk");
  var rows = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: "Produk tidak ditemukan." };
  }
  
  sheet.deleteRow(rowIndex);
  return { success: true, message: "Produk berhasil dihapus dari database." };
}

function updateOrderStatus(ss, orderId, status, kurir, resi) {
  var sheet = ss.getSheetByName("Transaksi");
  var rows = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString() === orderId.toString()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: "Transaksi tidak ditemukan." };
  }
  
  sheet.getRange(rowIndex, 11).setValue(status);
  
  if (status.toUpperCase() === "DIKIRIM") {
    if (kurir) sheet.getRange(rowIndex, 13).setValue(kurir);
    if (resi) sheet.getRange(rowIndex, 14).setValue(resi);
  }
  
  return { success: true, message: "Status pesanan berhasil diperbarui menjadi: " + status };
}

function updateSettings(ss, data) {
  var sheet = ss.getSheetByName("Pengaturan");
  var rows = sheet.getDataRange().getValues();
  var folder = getTargetFolder("Foto Produk - PANGG.CASUALS");
  
  for (var key in data) {
    var found = false;
    var rawValue = data[key];
    
    if (key === "Logo" || key === "Banner" || key === "QRIS") {
      if (rawValue && rawValue.startsWith("data:")) {
        rawValue = saveImageToDriveWithFolder(rawValue, "SETTING_" + key, "CONFIG", folder);
      }
    }
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString().trim() === key.toString().trim()) {
        sheet.getRange(i + 1, 2).setValue(rawValue);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, rawValue]);
    }
  }
  return { success: true, message: "Pengaturan toko berhasil diperbarui." };
}

function getSystemDiagnosticsGAS() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      success: true,
      sheetName: ss.getName(),
      sheetId: ss.getId(),
      sheetUrl: ss.getUrl()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getSettingsGAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { success: true, data: getSettings(ss) };
}

function getProductsAdminGAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { success: true, data: getProducts(ss, true) };
}

function getOrdersGAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { success: true, data: getOrders(ss) };
}

function verifyLoginGAS(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settings = getSettings(ss);
    var dbUser = settings.AdminUser || "admin";
    var dbPass = settings.AdminPass || "12345";
    
    if (username.trim() === dbUser.toString().trim() && password.trim() === dbPass.toString().trim()) {
      return { success: true, message: "Login sukses!" };
    } else {
      return { success: false, message: "Username atau password salah!" };
    }
  } catch (e) {
    return { success: false, message: "Koneksi database bermasalah: " + e.toString() };
  }
}

function getDashboardStatsGAS() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return { success: true, data: getDashboardStats(ss) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function createProductGAS(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return createProduct(ss, data);
}

function updateProductGAS(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return updateProduct(ss, data);
}

function deleteProductGAS(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return deleteProduct(ss, id);
}

function updateOrderStatusGAS(orderId, status, kurir, resi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return updateOrderStatus(ss, orderId, status, kurir, resi);
}

function updateSettingsGAS(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return updateSettings(ss, data);
}

function getTrackingDataByCode(ss, code) {
  var sheet = ss.getSheetByName("Transaksi");
  if (!sheet) return { success: false, message: "Tab database Transaksi tidak ditemukan." };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().trim(); });
  
  var codeIndex = headers.indexOf("TrackingCode");
  if (codeIndex === -1) {
    return { success: false, message: "Kolom pelacakan belum siap di database." };
  }
  
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (row[codeIndex].toString().trim().toUpperCase() === code.trim().toUpperCase()) {
      var item = {};
      for (var c = 0; c < headers.length; c++) {
        var headerName = headers[c];
        var val = row[c];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm:ss");
        } else if (val === null || val === undefined) {
          val = "";
        } else {
          val = val.toString();
        }
        item[headerName] = val;
      }
      return { success: true, data: item };
    }
  }
  return { success: false, message: "Tracking Code tidak terdaftar di database kami." };
}

function sanitizeForSerialization(obj) {
  if (obj === null || obj === undefined) return obj;
  if (obj instanceof Date) {
    return Utilities.formatDate(obj, "GMT+7", "yyyy-MM-dd HH:mm:ss");
  }
  if (Array.isArray(obj)) {
    return obj.map(sanitizeForSerialization);
  }
  if (typeof obj === "object") {
    var sanitizedObj = {};
    for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
        sanitizedObj[key] = sanitizeForSerialization(obj[key]);
      }
    }
    return sanitizedObj;
  }
  return obj;
}
