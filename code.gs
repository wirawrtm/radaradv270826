/**
 * GOOGLE APPS SCRIPT - code.gs
 * 
 * SCRIPT UNTUK MAKSIMALKAN PENAMBAHAN, UPDATE, DAN PENGHAPUSAN PARTNER
 * LANGSUNG DARI GOOGLE SPREADSHEET APPS SCRIPT.
 * 
 * Petunjuk Penggunaan:
 * 1. Buka Google Spreadsheet Anda.
 * 2. Klik menu "Ekstensi" > "Apps Script" (Extensions > Apps Script).
 * 3. Hapus kode lama atau tambahkan fungsi-fungsi di bawah ini ke dalam editor Apps Script Anda.
 * 4. Klik ikon Simpan (Save).
 * 5. Klik "Terapkan" > "Penerapan baru" (Deploy > New deployment).
 * 6. Pilih jenis penerapan: "Aplikasi Web" (Web app).
 * 7. Konfigurasikan:
 *    - Jalankan sebagai: "Saya" (Me)
 *    - Siapa yang memiliki akses: "Siapa saja" (Anyone)
 * 8. Klik "Terapkan" (Deploy), setujui izin jika diminta, lalu salin URL Aplikasi Web yang diberikan.
 */

// Helper: Membersihkan string agar pencocokan nama partner lebih akurat (case-insensitive & mengabaikan spasi/simbol)
function cleanForMatch(val) {
  if (!val) return "";
  return String(val)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

/**
 * Handle POST Requests
 */
function doPost(e) {
  var result = {};
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else {
      throw new Error("Tidak ada data postData yang diterima.");
    }
    
    var action = body.action;
    
    if (action === "addPartner") {
      result = handleAddPartner(body);
    } else if (action === "updatePartner") {
      result = handleUpdatePartner(body);
    } else if (action === "deletePartner") {
      result = handleDeletePartner(body);
    } else if (action === "batchActivity") {
      result = handleBatchActivity(body);
    } else if (action === "consolidateDatabase") {
      result = handleConsolidateDatabase(body);
    } else if (action === "updateEmployee") {
      result = handleUpdateEmployee(body);
    } else if (action === "deleteEmployee") {
      result = handleDeleteEmployee(body);
    } else if (action === "saveAccessRules") {
      result = handleSaveAccessRules(body);
    } else if (action === "backfillEmptyCategories") {
      result = handleBackfillEmptyCategories(body);
    } else {
      result = { 
        status: "error", 
        message: "Aksi '" + action + "' belum diimplementasikan di Google Apps Script ini." 
      };
    }
  } catch (err) {
    result = { 
      status: "error", 
      message: "Gagal memproses request: " + err.toString() 
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle GET Requests - Routes read requests to the appropriate Google Sheets reader
 */
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var user = e && e.parameter && e.parameter.user || "";
  var lot = e && e.parameter && e.parameter.lot || "";
  
  var result = {};
  
  try {
    if (!action) {
      result = {
        status: "success",
        message: "Google Apps Script aktif dan terhubung dengan sukses!"
      };
    } else if (action === "getWorkingData") {
      result = handleGetWorkingData(user);
    } else if (action === "getChannels") {
      result = handleGetChannels(user);
    } else if (action === "getDrSalesData") {
      result = handleGetDrSalesData(user);
    } else if (action === "getLotInfo") {
      result = handleGetLotInfo(lot);
    } else if (action === "getUserProfile") {
      result = handleGetUserProfile(user);
    } else if (action === "getEmployees") {
      result = handleGetEmployees();
    } else if (action === "getInitialData") {
      result = handleGetInitialData(user);
    } else if (action === "getAccessRules") {
      result = handleGetAccessRules();
    } else {
      result = {
        status: "error",
        message: "Aksi '" + action + "' belum diimplementasikan di Google Apps Script ini."
      };
    }
  } catch (err) {
    result = {
      status: "error",
      message: "Gagal memproses request GET: " + err.toString()
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Tambah Partner Baru (Atau update jika nama partner sudah ada)
 */
function handleAddPartner(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("channel");
  if (!sheet) {
    return { status: "error", message: "Sheet 'channel' tidak ditemukan." };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxPic = -1;
  var idxChannel = -1;
  var idxCat = -1;
  var idxProvince = -1;
  var idxArea = -1;
  var idxGroup = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (/pic|user|nama|analyst|solution/i.test(h)) idxPic = i;
    if (/channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(h)) idxChannel = i;
    if (/kategori|category|klasifikasi|^cat$/i.test(h)) idxCat = i;
    if (/provinsi|province/i.test(h)) idxProvince = i;
    if (/^area$/i.test(h)) idxArea = i;
    if (/group|tim|divisi|division/i.test(h)) idxGroup = i;
  }
  
  if (idxChannel === -1) {
    return { status: "error", message: "Kolom nama partner tidak ditemukan di header sheet 'channel'." };
  }
  
  // Mencegah duplikasi: Jika partner sudah ada, langsung perbarui baris yang ada
  if (body.name) {
    var cleanNewName = cleanForMatch(body.name);
    var existingIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if (cleanForMatch(data[r][idxChannel]) === cleanNewName) {
        existingIndex = r;
        break;
      }
    }
    
    if (existingIndex !== -1) {
      var userProvince = body.province || "";
      var userArea = "";
      
      // Cari data wilayah berdasarkan kecocokan PIC lain
      if (body.pic && data.length > 1) {
        var cleanPic = cleanForMatch(body.pic);
        for (var r = 1; r < data.length; r++) {
          if (r !== existingIndex && idxPic !== -1 && cleanForMatch(data[r][idxPic]) === cleanPic) {
            if (idxProvince !== -1 && data[r][idxProvince]) {
              userProvince = String(data[r][idxProvince]).trim();
            }
            if (idxArea !== -1 && data[r][idxArea]) {
              userArea = String(data[r][idxArea]).trim();
            }
            break;
          }
        }
      }
      
      // Update data sel partner yang ada
      if (idxPic !== -1 && body.pic !== undefined) {
        sheet.getRange(existingIndex + 1, idxPic + 1).setValue(body.pic);
      }
      if (idxProvince !== -1 && userProvince) {
        sheet.getRange(existingIndex + 1, idxProvince + 1).setValue(userProvince);
      }
      if (idxArea !== -1 && userArea) {
        sheet.getRange(existingIndex + 1, idxArea + 1).setValue(userArea);
      }
      if (idxCat !== -1 && body.category !== undefined && body.category !== "") {
        sheet.getRange(existingIndex + 1, idxCat + 1).setValue(body.category);
      }
      if (idxGroup !== -1 && body.group !== undefined && body.group !== "") {
        sheet.getRange(existingIndex + 1, idxGroup + 1).setValue(body.group);
      }
      
      return {
        status: "success",
        id: existingIndex + 1,
        message: "Partner '" + body.name + "' sudah ada di database, data berhasil diperbarui."
      };
    }
  }
  
  // Mencari wilayah berdasarkan PIC jika menambahkan partner baru
  var userProvince = body.province || "";
  var userArea = "";
  
  if (body.pic && data.length > 1) {
    var cleanPic = cleanForMatch(body.pic);
    for (var r = 1; r < data.length; r++) {
      if (idxPic !== -1 && cleanForMatch(data[r][idxPic]) === cleanPic) {
        if (idxProvince !== -1 && data[r][idxProvince]) {
          userProvince = String(data[r][idxProvince]).trim();
        }
        if (idxArea !== -1 && data[r][idxArea]) {
          userArea = String(data[r][idxArea]).trim();
        }
        break;
      }
    }
  }
  
  // Buat baris baru sesuai urutan header
  var newRow = [];
  for (var i = 0; i < headers.length; i++) {
    newRow.push("");
  }
  
  if (idxChannel !== -1) newRow[idxChannel] = body.name || "";
  if (idxCat !== -1) newRow[idxCat] = body.category || "";
  if (idxPic !== -1) newRow[idxPic] = body.pic || "";
  if (idxProvince !== -1) newRow[idxProvince] = userProvince || "";
  if (idxArea !== -1) newRow[idxArea] = userArea || "";
  if (idxGroup !== -1) newRow[idxGroup] = body.group || "";
  
  sheet.appendRow(newRow);
  
  var newId = data.length + 1; // ID baris baru di spreadsheet
  return {
    status: "success",
    id: newId,
    message: "Partner '" + body.name + "' berhasil ditambahkan!"
  };
}

/**
 * Edit / Update Partner
 */
function handleUpdatePartner(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("channel");
  if (!sheet) {
    return { status: "error", message: "Sheet 'channel' tidak ditemukan." };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxPic = -1;
  var idxChannel = -1;
  var idxCat = -1;
  var idxProvince = -1;
  var idxArea = -1;
  var idxGroup = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (/pic|user|nama|analyst|solution/i.test(h)) idxPic = i;
    if (/channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(h)) idxChannel = i;
    if (/kategori|category|klasifikasi|^cat$/i.test(h)) idxCat = i;
    if (/provinsi|province/i.test(h)) idxProvince = i;
    if (/^area$/i.test(h)) idxArea = i;
    if (/group|tim|divisi|division/i.test(h)) idxGroup = i;
  }
  
  var rowIndex = -1;
  var rowNum = Number(body.id);
  
  // 1. Validasi via Row Number (ID baris)
  if (!isNaN(rowNum) && rowNum > 1 && rowNum <= data.length) {
    var potentialRow = data[rowNum - 1];
    if (potentialRow && idxChannel !== -1) {
      var currentClean = cleanForMatch(potentialRow[idxChannel]);
      var cleanOrig = body.originalName ? cleanForMatch(body.originalName) : "";
      var cleanName = body.name ? cleanForMatch(body.name) : "";
      
      if (cleanOrig && currentClean === cleanOrig) {
        rowIndex = rowNum - 1;
      } else if (cleanName && currentClean === cleanName) {
        rowIndex = rowNum - 1;
      }
    }
  }
  
  // 2. Fallback pencarian berdasarkan nama asli
  if (rowIndex === -1 && body.originalName && idxChannel !== -1) {
    var targetClean = cleanForMatch(body.originalName);
    for (var r = 1; r < data.length; r++) {
      if (cleanForMatch(data[r][idxChannel]) === targetClean) {
        rowIndex = r;
        break;
      }
    }
  }
  
  // 3. Fallback pencarian berdasarkan nama baru
  if (rowIndex === -1 && body.name && idxChannel !== -1) {
    var targetClean = cleanForMatch(body.name);
    for (var r = 1; r < data.length; r++) {
      if (cleanForMatch(data[r][idxChannel]) === targetClean) {
        rowIndex = r;
        break;
      }
    }
  }
  
  if (rowIndex > 0 && rowIndex < data.length) {
    var userProvince = body.province || "";
    var userArea = "";
    
    // Tarik data wilayah berdasarkan PIC
    if (body.pic && data.length > 1) {
      var cleanPic = cleanForMatch(body.pic);
      for (var r = 1; r < data.length; r++) {
        if (r !== rowIndex && idxPic !== -1 && cleanForMatch(data[r][idxPic]) === cleanPic) {
          if (idxProvince !== -1 && data[r][idxProvince]) {
            userProvince = String(data[r][idxProvince]).trim();
          }
          if (idxArea !== -1 && data[r][idxArea]) {
            userArea = String(data[r][idxArea]).trim();
          }
          break;
        }
      }
    }
    
    // Simpan perubahan ke baris partner
    if (idxPic !== -1 && body.pic !== undefined) {
      sheet.getRange(rowIndex + 1, idxPic + 1).setValue(body.pic);
    }
    if (idxProvince !== -1) {
      var nextProvince = userProvince || data[rowIndex][idxProvince] || "";
      sheet.getRange(rowIndex + 1, idxProvince + 1).setValue(nextProvince);
    }
    if (idxArea !== -1) {
      var nextArea = userArea || data[rowIndex][idxArea] || "";
      sheet.getRange(rowIndex + 1, idxArea + 1).setValue(nextArea);
    }
    if (idxChannel !== -1 && body.name !== undefined && body.name !== "") {
      sheet.getRange(rowIndex + 1, idxChannel + 1).setValue(body.name);
    }
    if (idxCat !== -1 && body.category !== undefined && body.category !== "") {
      sheet.getRange(rowIndex + 1, idxCat + 1).setValue(body.category);
    }
    if (idxGroup !== -1 && body.group !== undefined && body.group !== "") {
      sheet.getRange(rowIndex + 1, idxGroup + 1).setValue(body.group);
    }
    
    return { 
      status: "success", 
      message: "Partner '" + body.name + "' berhasil diperbarui!" 
    };
  } else {
    // Jika tidak ditemukan sama sekali, tambahkan partner baru
    return handleAddPartner(body);
  }
}

/**
 * Hapus Partner
 */
function handleDeletePartner(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("channel");
  if (!sheet) {
    return { status: "error", message: "Sheet 'channel' tidak ditemukan." };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxPic = -1;
  var idxChannel = -1;
  var idxCat = -1;
  var idxProvince = -1;
  var idxArea = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (/pic|user|nama|analyst|solution/i.test(h)) idxPic = i;
    if (/channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(h)) idxChannel = i;
    if (/kategori|category|klasifikasi|^cat$/i.test(h)) idxCat = i;
    if (/provinsi|province/i.test(h)) idxProvince = i;
    if (/^area$/i.test(h)) idxArea = i;
  }
  
  if (idxChannel === -1) {
    return { status: "error", message: "Kolom nama partner tidak ditemukan di sheet." };
  }
  
  var rowIndex = -1;
  var rowNum = Number(body.id);
  
  // 1. Check if rowNum is within bounds and the name actually matches using cleanForMatch
  if (!isNaN(rowNum) && rowNum > 1 && rowNum <= data.length) {
    var potentialRow = data[rowNum - 1];
    if (potentialRow && idxChannel !== -1 && body.name) {
      var currentClean = cleanForMatch(potentialRow[idxChannel]);
      var cleanName = cleanForMatch(body.name);
      if (currentClean === cleanName) {
        rowIndex = rowNum - 1;
      }
    }
  }
  
  // 2. Fallback to searching by name across all rows using cleanForMatch
  if (rowIndex === -1 && body.name && idxChannel !== -1) {
    var targetClean = cleanForMatch(body.name);
    for (var r = 1; r < data.length; r++) {
      if (cleanForMatch(data[r][idxChannel]) === targetClean) {
        rowIndex = r;
        break;
      }
    }
  }
  
  if (rowIndex > 0 && rowIndex < data.length) {
    var deletedName = data[rowIndex][idxChannel];
    sheet.deleteRow(rowIndex + 1); // deleteRow menggunakan index berbasis 1
    return { 
      status: "success", 
      message: "Partner '" + deletedName + "' berhasil dihapus!" 
      };
  }
  
  return {
    status: "error",
    message: rowIndex === -1
      ? "Data partner '" + (body.name || "") + "' tidak ditemukan."
      : "Indeks baris tidak valid untuk penghapusan."
  };
}

/**
 * Helper: Ambil semua data sel dari sheet berdasarkan nama
 */
function getSheetValues(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  return sheet.getDataRange().getValues();
}

/**
 * Handler: Ambil data Working/Checker
 */
function handleGetWorkingData(user) {
  var data = getSheetValues("working");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  var headers = data[0];
  
  function getIdx(patterns) {
    return headers.findIndex(function(h) {
      return patterns.test(String(h).trim());
    });
  }
  
  var idx = {
    lot: getIdx(/^lot package$|^lot$/i),
    hybrid: getIdx(/^hybrid$|^material$/i),
    stock: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
    aging: getIdx(/^aging \(month\)/i),
    exp: getIdx(/^exp date$|^expired$/i),
    kiosk: getIdx(/^channel$|^kiosk$/i),
    cat: getIdx(/category|kategori|klasifikasi|^cat/i),
    crops: getIdx(/^crops$/i),
    time: getIdx(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i),
    cond: getIdx(/^condition$|^kondisi$/i),
    dr: getIdx(/^shipping date$|^dr date$/i),
    user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
    pog: getIdx(/^pog$|^selisih$/i)
  };
  
  var monthIndices = getMonthIndices(headers);
  var updMonthIndices = getUpdMonthIndices(headers);
  
  var result = data.slice(1).filter(function(row) {
    var kioskVal = idx.kiosk !== -1 ? row[idx.kiosk] : undefined;
    var userVal = idx.user !== -1 ? row[idx.user] : undefined;
    var lotVal = idx.lot !== -1 ? row[idx.lot] : undefined;
    return (
      row.length > 0 &&
      ((kioskVal !== undefined && kioskVal !== "") ||
       (userVal !== undefined && userVal !== "") ||
       (lotVal !== undefined && lotVal !== ""))
    );
  }).map(function(row) {
    var rowItem = {
      lot: idx.lot !== -1 ? row[idx.lot] : "",
      hybrid: idx.hybrid !== -1 ? row[idx.hybrid] : "",
      crops: (idx.crops !== -1 && row[idx.crops] !== "" && row[idx.crops] !== undefined) ? row[idx.crops] : "Uncategorized Crops",
      stock: (idx.stock !== -1 && row[idx.stock] !== "" && row[idx.stock] !== undefined) ? Number(row[idx.stock]) || 0 : 0,
      aging: (idx.aging !== -1 && row[idx.aging] !== "" && row[idx.aging] !== undefined) ? row[idx.aging] : "-",
      expired: (idx.exp !== -1 && row[idx.exp]) ? formatMyDate(row[idx.exp]) : "N/A",
      drDate: (idx.dr !== -1 && row[idx.dr]) ? formatMyDate(row[idx.dr]) : "N/A",
      kiosk: idx.kiosk !== -1 ? row[idx.kiosk] : "",
      category: (idx.cat !== -1 && row[idx.cat] !== "" && row[idx.cat] !== undefined) ? String(row[idx.cat]).trim() : "",
      timestamp: (idx.time !== -1 && row[idx.time]) ? row[idx.time] : "",
      condition: (idx.cond !== -1 && row[idx.cond]) ? row[idx.cond] : "tetap",
      user: (idx.user !== -1 && row[idx.user]) ? String(row[idx.user]).trim() : "",
      pog: (idx.pog !== -1 && row[idx.pog] !== "" && row[idx.pog] !== undefined) ? Number(row[idx.pog]) || 0 : 0
    };
    
    var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des"];
    months.forEach(function(m, mIdx) {
      var colIdx = monthIndices[mIdx];
      if (colIdx !== -1 && colIdx < row.length) {
        rowItem[m.toLowerCase()] = (row[colIdx] !== "" && row[colIdx] !== undefined) ? Number(row[colIdx]) || 0 : 0;
      } else {
        rowItem[m.toLowerCase()] = 0;
      }
      
      var updColIdx = updMonthIndices[mIdx];
      if (updColIdx !== -1 && updColIdx < row.length) {
        rowItem["upd_" + m.toLowerCase()] = (row[updColIdx] !== "" && row[updColIdx] !== undefined) ? String(row[updColIdx]).trim() : "";
      } else {
        rowItem["upd_" + m.toLowerCase()] = "";
      }
    });
    return rowItem;
  });
  
  return { status: "success", data: result };
}

/**
 * Helper: Ambil index kolom bulan
 */
function getMonthIndices(headers) {
  var synonyms = [
    ["jan", "januari", "january"],
    ["feb", "februari", "february"],
    ["mar", "maret", "march"],
    ["apr", "april"],
    ["mei", "may"],
    ["jun", "juni", "june"],
    ["jul", "juli", "july"],
    ["ags", "agu", "agst", "agustus", "aug", "august"],
    ["sep", "sept", "september"],
    ["okt", "oct", "oktober", "october"],
    ["nov", "november"],
    ["des", "dec", "desember", "december"]
  ];
  var matchedIndices = [];
  for (var m = 0; m < 12; m++) {
    var list = synonyms[m];
    var idx = headers.findIndex(function(h) {
      var hStr = String(h || "").trim().toLowerCase();
      return list.some(function(syn) { return hStr === syn; });
    });
    matchedIndices.push(idx);
  }
  return matchedIndices;
}

/**
 * Helper: Ambil index kolom update bulan
 */
function getUpdMonthIndices(headers) {
  var synonyms = [
    ["jan", "januari", "january"],
    ["feb", "februari", "february"],
    ["mar", "maret", "march"],
    ["apr", "april"],
    ["mei", "may"],
    ["jun", "juni", "june"],
    ["jul", "juli", "july"],
    ["ags", "agu", "agst", "agustus", "aug", "august"],
    ["sep", "sept", "september"],
    ["okt", "oct", "oktober", "october"],
    ["nov", "november"],
    ["des", "dec", "desember", "december"]
  ];
  var matchedIndices = [];
  for (var m = 0; m < 12; m++) {
    var list = synonyms[m];
    var idx = headers.findIndex(function(h) {
      if (h === undefined || h === null) return false;
      var hStr = String(h).trim().toLowerCase().replace(/[\s_\-\/]/g, "");
      if (!hStr.startsWith("upd")) return false;
      var remains = hStr.substring(3);
      return list.some(function(syn) { return remains === syn; });
    });
    matchedIndices.push(idx);
  }
  return matchedIndices;
}

/**
 * Handler: Ambil data Partner/Channel yang sesuai hak akses
 */
function handleGetChannels(user) {
  var data = getSheetValues("channel");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  var lowerUser = String(user).trim().toLowerCase();
  
  var empData = getSheetValues("employee");
  var authorizedPICs = [lowerUser];
  var isBusinessAnalyst = 
    lowerUser === "adityawiratama" || 
    lowerUser.includes("adityawiratama") || 
    lowerUser === "analyst" || 
    lowerUser === "businessanalyst";
    
  var empHeaders = empData ? empData[0] : null;
  var matchedRow = null;
  var idxE = { name: -1, email: -1, upline: -1, pos: -1 };
  
  if (empData && empData.length > 0) {
    idxE.name = empHeaders.findIndex(function(h) { return /nama|name|pic/i.test(String(h).trim()); });
    idxE.email = empHeaders.findIndex(function(h) { return /email|user/i.test(String(h).trim()); });
    idxE.upline = empHeaders.findIndex(function(h) { return /upline|spv|supervisor|atasan|manager/i.test(String(h).trim()); });
    idxE.pos = empHeaders.findIndex(function(h) { return /position|jabatan/i.test(String(h).trim()); });
    
    matchedRow = findEmployeeRow(user, empData);
    var userAliases = new Set([lowerUser]);
    
    if (matchedRow) {
      var emailIdx = empHeaders.findIndex(function(h) { return /email/i.test(String(h).trim()); });
      var userIdx = empHeaders.findIndex(function(h) { return /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()); });
      var rowName = idxE.name !== -1 ? String(matchedRow[idxE.name] || "").trim().toLowerCase() : "";
      var rowEmail = emailIdx !== -1 ? String(matchedRow[emailIdx] || "").trim().toLowerCase() : "";
      var rowUser = userIdx !== -1 ? String(matchedRow[userIdx] || "").trim().toLowerCase() : "";
      var rowPos = idxE.pos !== -1 ? String(matchedRow[idxE.pos] || "").trim().toLowerCase() : "";
      
      if (rowName !== "") userAliases.add(rowName);
      if (rowEmail !== "") userAliases.add(rowEmail);
      if (rowUser !== "") userAliases.add(rowUser);
      
      var cleanRowPos = rowPos.replace(/\s+/g, "");
      var levelIdx = empHeaders.findIndex(function(h) { return /level|grade/i.test(String(h).trim()); });
      var rowLevel = levelIdx !== -1 ? String(matchedRow[levelIdx] || "").trim().toLowerCase() : "";
      if (cleanRowPos === "businessanalyst" || cleanRowPos === "analyst" || rowLevel === "admin") {
        isBusinessAnalyst = true;
      }
    }
    
    if (isBusinessAnalyst) {
      var emailIdx = empHeaders.findIndex(function(h) { return /email/i.test(String(h).trim()); });
      var userIdx = empHeaders.findIndex(function(h) { return /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()); });
      var groupIdx = empHeaders.findIndex(function(h) { return /group|tim|divisi|division/i.test(String(h).trim()); });
      var levelIdx = empHeaders.findIndex(function(h) { return /level|grade/i.test(String(h).trim()); });
      
      var adminGroup = "";
      if (matchedRow && levelIdx !== -1 && groupIdx !== -1) {
        var rowLevel2 = String(matchedRow[levelIdx] || "").trim().toLowerCase();
        if (rowLevel2 === "admin") {
          adminGroup = String(matchedRow[groupIdx] || "").trim().toLowerCase();
        }
      }
      
      empData.slice(1).forEach(function(row) {
        var rowGroup = groupIdx !== -1 ? String(row[groupIdx] || "").trim().toLowerCase() : "";
        var rowName = idxE.name !== -1 ? String(row[idxE.name] || "").trim().toLowerCase() : "";
        var rowEmail = emailIdx !== -1 ? String(row[emailIdx] || "").trim().toLowerCase() : "";
        var rowUser = userIdx !== -1 ? String(row[userIdx] || "").trim().toLowerCase() : "";
        
        if (adminGroup && adminGroup !== "all" && adminGroup !== "") {
          if (rowGroup !== adminGroup && rowName !== lowerUser && rowEmail !== lowerUser && rowUser !== lowerUser) {
            return;
          }
        }
        
        if (rowName !== "") userAliases.add(rowName);
        if (rowEmail !== "") userAliases.add(rowEmail);
        if (rowUser !== "") userAliases.add(rowUser);
      });
    }
    
    var queue = Array.from(userAliases);
    var visited = new Set(queue);
    userAliases.forEach(function(alias) {
      if (!authorizedPICs.includes(alias)) authorizedPICs.push(alias);
    });
    
    while (queue.length > 0) {
      var currentUpline = queue.shift();
      empData.slice(1).forEach(function(row) {
        var empNameRaw = idxE.name !== -1 ? String(row[idxE.name] || "").trim() : "";
        var empNameLower = empNameRaw.toLowerCase();
        var empEmailRaw = idxE.email !== -1 ? String(row[idxE.email] || "").trim() : "";
        var empEmailLower = empEmailRaw.toLowerCase();
        var empUplineRaw = idxE.upline !== -1 ? String(row[idxE.upline] || "").trim() : "";
        var empUplineLower = empUplineRaw.toLowerCase();
        
        if (empUplineLower !== "") {
          var isMatch = empUplineLower === currentUpline || empUplineLower.includes(currentUpline) || currentUpline.includes(empUplineLower);
          if (isMatch) {
            if (empNameLower !== "" && !visited.has(empNameLower)) {
              visited.add(empNameLower);
              queue.push(empNameLower);
              if (!authorizedPICs.includes(empNameLower)) authorizedPICs.push(empNameLower);
            }
            if (empEmailLower !== "" && !visited.has(empEmailLower)) {
              visited.add(empEmailLower);
              queue.push(empEmailLower);
              if (!authorizedPICs.includes(empEmailLower)) authorizedPICs.push(empEmailLower);
            }
          }
        }
      });
    }
  }
  
  var headers = data[0];
  var idx = {
    pic: headers.findIndex(function(h) { return /pic|user|nama|analyst|solution/i.test(String(h).trim()); }),
    channel: headers.findIndex(function(h) { return /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()); }),
    cat: headers.findIndex(function(h) { return /kategori|category|klasifikasi|^cat$/i.test(String(h).trim()); }),
    upline: headers.findIndex(function(h) { return /upline|spv|supervisor/i.test(String(h).trim()); }),
    area: headers.findIndex(function(h) { return /area|provinsi|province|wilayah/i.test(String(h).trim()); }),
    group: headers.findIndex(function(h) { return /group|tim|divisi|division/i.test(String(h).trim()); })
  };
  
  var picToAreaMap = {};
  if (empData && empData.length > 1) {
    var empHeadersVal = empData[0];
    var nameCol = empHeadersVal.findIndex(function(h) { return /nama|name|pic/i.test(String(h).trim()); });
    var emailCol = empHeadersVal.findIndex(function(h) { return /email|user/i.test(String(h).trim()); });
    var areaCol = empHeadersVal.findIndex(function(h) { return /area/i.test(String(h).trim()); });
    var provCol = empHeadersVal.findIndex(function(h) { return /province|provinsi/i.test(String(h).trim()); });
    
    empData.slice(1).forEach(function(empRow) {
      var empName = nameCol !== -1 ? String(empRow[nameCol] || "").trim().toLowerCase() : "";
      var empEmail = emailCol !== -1 ? String(empRow[emailCol] || "").trim().toLowerCase() : "";
      var empArea = (areaCol !== -1 && empRow[areaCol] !== "" && empRow[areaCol] !== undefined) ? String(empRow[areaCol]).trim() :
                    (provCol !== -1 && empRow[provCol] !== "" && empRow[provCol] !== undefined) ? String(empRow[provCol]).trim() : "";
      if (empArea) {
        if (empName) picToAreaMap[empName] = empArea;
        if (empEmail) picToAreaMap[empEmail] = empArea;
      }
    });
  }
  
  var channels = data.slice(1).map(function(row, i) {
    if (row[0] === "" || row[0] === undefined) return null;
    var catValue = (idx.cat !== -1 && row[idx.cat] !== "" && row[idx.cat] !== undefined) ? String(row[idx.cat]).trim() : "Uncategorized";
    var rowGroup = (idx.group !== -1 && row[idx.group] !== "" && row[idx.group] !== undefined) ? String(row[idx.group]).trim() : "";
    
    if (idx.pic !== -1 && idx.channel !== -1) {
      var picLower = String(row[idx.pic] || "").trim().toLowerCase();
      var uplineLower = idx.upline !== -1 ? String(row[idx.upline] || "").trim().toLowerCase() : "";
      
      var isAuth = isBusinessAnalyst || picLower === "" || picLower === "tanpa pic" || picLower === "tidak ada" || picLower === lowerUser ||
                   (lowerUser !== "" && picLower.includes(lowerUser)) || uplineLower === lowerUser || (lowerUser !== "" && uplineLower.includes(lowerUser)) ||
                   authorizedPICs.some(function(auth) { return picLower === auth || (auth !== "" && picLower.includes(auth)); });
                   
      if (isAuth) {
        var sheetArea = (idx.area !== -1 && row[idx.area] !== "" && row[idx.area] !== undefined) ? String(row[idx.area]).trim() : "";
        var resolvedArea = sheetArea || picToAreaMap[picLower] || "-";
        return {
          id: i + 2,
          name: row[idx.channel],
          category: catValue,
          pic: String(row[idx.pic] || "").trim(),
          upline: idx.upline !== -1 ? String(row[idx.upline] || "").trim() : "",
          area: resolvedArea,
          group: rowGroup
        };
      }
    } else if (idx.channel !== -1) {
      var sheetArea = (idx.area !== -1 && row[idx.area] !== "" && row[idx.area] !== undefined) ? String(row[idx.area]).trim() : "";
      return {
        id: i + 2,
        name: row[idx.channel],
        category: catValue,
        pic: "",
        upline: "",
        area: sheetArea || "-",
        group: rowGroup
      };
    }
    return null;
  }).filter(Boolean);
  
  return { status: "success", data: channels };
}

/**
 * Handler: Ambil data DR Sales
 */
function handleGetDrSalesData(user) {
  var hybridMap = {};
  var hData = getSheetValues("hybrid");
  if (hData && hData.length > 1) {
    var hHeaders = hData[0];
    var idxH = {
      desc: hHeaders.findIndex(function(h) { return /material.*desc|description/i.test(String(h).trim()); }),
      hybrid: hHeaders.findIndex(function(h) { return /^hybrid$/i.test(String(h).trim()); }),
      crops: hHeaders.findIndex(function(h) { return /^crops$/i.test(String(h).trim()); })
    };
    if (idxH.desc !== -1 && idxH.hybrid !== -1) {
      hData.slice(1).forEach(function(row) {
        var mDesc = String(row[idxH.desc] || "").trim().toLowerCase();
        if (mDesc) {
          hybridMap[mDesc] = {
            hybrid: String(row[idxH.hybrid] || "").trim(),
            crops: (idxH.crops !== -1 && row[idxH.crops] !== undefined) ? String(row[idxH.crops]).trim() : ""
          };
        }
      });
    }
  }

  var data = getSheetValues("dr");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  var headers = data[0];

  var hIdx = {
    qty: headers.findIndex(function(h) { return /qty|quantity/i.test(String(h).trim()); }),
    type: headers.findIndex(function(h) { return /order type/i.test(String(h).trim()); }),
    channel: headers.findIndex(function(h) { return /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()); }),
    lot: headers.findIndex(function(h) { return /lot/i.test(String(h).trim()); }),
    desc: headers.findIndex(function(h) { return /material.*desc|description/i.test(String(h).trim()); }),
    dr: headers.findIndex(function(h) { return /dr date|shipping date/i.test(String(h).trim()); }),
    exp: headers.findIndex(function(h) { return /exp date|expired/i.test(String(h).trim()); })
  };

  var result = data.slice(1).filter(function(row) {
    return hIdx.type !== -1 && String(row[hIdx.type] || "").trim().toLowerCase() === "sales";
  }).map(function(row) {
    var rawDesc = hIdx.desc !== -1 ? String(row[hIdx.desc] || "").trim() : "";
    var mapInfo = hybridMap[rawDesc.toLowerCase()] || { hybrid: rawDesc, crops: "" };
    var drValue = (hIdx.dr !== -1 && row[hIdx.dr]) ? formatMyDate(row[hIdx.dr]) : "N/A";
    var expValue = (hIdx.exp !== -1 && row[hIdx.exp]) ? formatMyDate(row[hIdx.exp]) : "N/A";
    return {
      lot: hIdx.lot !== -1 ? String(row[hIdx.lot] || "").trim().toUpperCase() : "",
      hybrid: mapInfo.hybrid,
      crops: mapInfo.crops,
      channel: hIdx.channel !== -1 ? String(row[hIdx.channel] || "").trim() : "",
      qty: hIdx.qty !== -1 ? Number(row[hIdx.qty]) || 0 : 0,
      drDate: drValue,
      expired: expValue
    };
  });
  return { status: "success", data: result };
}

/**
 * Handler: Ambil info lot tertentu
 */
function handleGetLotInfo(lotNo) {
  if (!lotNo) return { status: "error", message: "Lot number is required" };
  var hybridMap = {};
  var hData = getSheetValues("hybrid");
  if (hData && hData.length > 1) {
    var hHeaders = hData[0];
    var idxH = {
      desc: hHeaders.findIndex(function(h) { return /material.*desc|description/i.test(String(h).trim()); }),
      hybrid: hHeaders.findIndex(function(h) { return /^hybrid$/i.test(String(h).trim()); }),
      crops: hHeaders.findIndex(function(h) { return /^crops$/i.test(String(h).trim()); })
    };
    if (idxH.desc !== -1 && idxH.hybrid !== -1) {
      hData.slice(1).forEach(function(row) {
        var mDesc = String(row[idxH.desc] || "").trim().toLowerCase();
        if (mDesc) {
          hybridMap[mDesc] = {
            hybrid: String(row[idxH.hybrid] || "").trim(),
            crops: (idxH.crops !== -1 && row[idxH.crops] !== undefined) ? String(row[idxH.crops]).trim() : ""
          };
        }
      });
    }
  }
  var data = getSheetValues("dr");
  if (!data || data.length <= 1) return { status: "error", message: "No data" };
  var headers = data[0];
  var idx = {
    lot: headers.findIndex(function(h) { return /lot/i.test(String(h).trim()); }),
    desc: headers.findIndex(function(h) { return /material.*desc|description/i.test(String(h).trim()); }),
    dr: headers.findIndex(function(h) { return /dr date|shipping date/i.test(String(h).trim()); }),
    exp: headers.findIndex(function(h) { return /exp date|expired/i.test(String(h).trim()); })
  };
  if (idx.lot === -1) return { status: "error", message: "Lot column missing" };
  
  function calcMonths(start, end) {
    try {
      if (!start || !end || start === "" || end === "") return "";
      var d1 = new Date(start);
      var d2 = new Date(end);
      if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return "";
      return Math.round((d2.getTime() - d1.getTime()) / (1000 * 3600 * 24) / 30.416);
    } catch (e) {
      return "";
    }
  }

  var todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);
  var targetLot = String(lotNo).trim().toUpperCase();
  var foundRow = data.slice(1).find(function(row) {
    return String(row[idx.lot] || "").trim().toUpperCase() === targetLot;
  });
  if (foundRow) {
    var rawDesc = idx.desc !== -1 ? String(foundRow[idx.desc] || "").trim() : "Unknown Material";
    var mapInfo = hybridMap[rawDesc.toLowerCase()] || { hybrid: rawDesc, crops: "" };
    var drDateVal = idx.dr !== -1 ? foundRow[idx.dr] : "";

    return {
      status: "success",
      data: {
        desc: mapInfo.hybrid,
        crops: mapInfo.crops,
        drDate: idx.dr !== -1 ? formatMyDate(drDateVal) : "N/A",
        expDate: idx.exp !== -1 ? formatMyDate(foundRow[idx.exp]) : "N/A",
        aging: calcMonths(drDateVal, todayDate)
      }
    };
  }
  return { status: "error", message: "Lot not found" };
}

/**
 * Handler: Ambil daftar karyawan/user
 */
function handleGetEmployees() {
  var data = getSheetValues("employee");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  var headers = data[0];
  var emailIdx = headers.findIndex(function(h) { return /email/i.test(String(h).trim()); });
  var userIdx = headers.findIndex(function(h) { return /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()); });
  var idx = {
    name: headers.findIndex(function(h) { return /nama|name|pic/i.test(String(h).trim()); }),
    email: emailIdx !== -1 ? emailIdx : headers.findIndex(function(h) { return /email|user/i.test(String(h).trim()); }),
    username: userIdx !== -1 ? userIdx : -1,
    pos: headers.findIndex(function(h) { return /position|jabatan/i.test(String(h).trim()); }),
    prov: headers.findIndex(function(h) { return /province|provinsi/i.test(String(h).trim()); }),
    area: headers.findIndex(function(h) { return /area/i.test(String(h).trim()); }),
    upline: headers.findIndex(function(h) { return /upline|spv|supervisor|atasan|manager/i.test(String(h).trim()); }),
    password: headers.findIndex(function(h) { return /password|pass/i.test(String(h).trim()); }),
    level: headers.findIndex(function(h) { return /level|grade/i.test(String(h).trim()); }),
    group: headers.findIndex(function(h) { return /group|tim|divisi|division/i.test(String(h).trim()); })
  };

  var result = data.slice(1).filter(function(row) {
    return row[idx.name] !== "" && row[idx.name] !== undefined;
  }).map(function(row) {
    var p = idx.pos !== -1 ? row[idx.pos] : "Business Solution";
    return {
      name: idx.name !== -1 ? String(row[idx.name] || "").trim() : "",
      email: idx.email !== -1 ? String(row[idx.email] || "").trim() : "",
      user: idx.username !== -1 ? String(row[idx.username] || "").trim() : "",
      position: normalizePosition(p),
      province: idx.prov !== -1 ? String(row[idx.prov] || "").trim() : "-",
      area: idx.area !== -1 ? String(row[idx.area] || "").trim() : "-",
      upline: idx.upline !== -1 ? String(row[idx.upline] || "").trim() : "",
      password: idx.password !== -1 ? String(row[idx.password] || "").trim() : "",
      level: (idx.level !== -1 && row[idx.level] !== "" && row[idx.level] !== undefined) ? row[idx.level] : null,
      group: idx.group !== -1 ? String(row[idx.group] || "").trim() : ""
    };
  });
  
  var seenEmployees = new Set();
  var dedupedResult = [];
  for (var i = 0; i < result.length; i++) {
    var emp = result[i];
    var cleanName = String(emp.name || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    if (!seenEmployees.has(cleanName) && cleanName !== "") {
      seenEmployees.add(cleanName);
      dedupedResult.push(emp);
    }
  }
  
  return { status: "success", data: dedupedResult };
}

/**
 * Handler: Ambil profil user login
 */
function handleGetUserProfile(user) {
  var employeesRes = handleGetEmployees();
  var employees = employeesRes.status === "success" ? employeesRes.data : [];
  if (employees.length === 0) {
    return { status: "error", message: "Data employee kosong atau tidak ditemukan" };
  }

  var lowerUser = String(user).trim().toLowerCase();

  var foundEmployee = employees.find(function(emp) {
    var rowUser = String(emp.user || "").trim().toLowerCase();
    var rowEmail = String(emp.email || "").trim().toLowerCase();
    var rowUserLocal = rowEmail.includes("@") ? rowEmail.split("@")[0] : rowEmail;

    if (rowUser !== "") {
      return rowUser === lowerUser || rowEmail === lowerUser;
    } else {
      return rowUserLocal === lowerUser || rowEmail === lowerUser;
    }
  });

  if (!foundEmployee) {
    return { status: "error", message: "Username tidak ditemukan: " + user };
  }

  var resolvedUser = foundEmployee.user || (foundEmployee.email ? (foundEmployee.email.includes("@") ? foundEmployee.email.split("@")[0] : foundEmployee.email) : "");

  var profile = {
    name: foundEmployee.name,
    email: foundEmployee.email,
    user: resolvedUser,
    position: foundEmployee.position,
    province: foundEmployee.province,
    area: foundEmployee.area,
    password: foundEmployee.password,
    upline: foundEmployee.upline,
    level: foundEmployee.level,
    group: foundEmployee.group,
    subordinates: []
  };

  var userAliases = new Set([lowerUser]);
  if (profile.name !== "") userAliases.add(profile.name.toLowerCase());
  if (profile.email !== "") userAliases.add(profile.email.toLowerCase());
  if (profile.user !== "") userAliases.add(profile.user.toLowerCase());

  var cleanProfilePos = profile.position ? profile.position.toLowerCase().replace(/\s+/g, "") : "";
  var profileLevelClean = profile.level ? String(profile.level).toLowerCase().trim() : "";
  var isBusinessAnalyst =
    lowerUser === "adityawiratama" ||
    lowerUser.includes("adityawiratama") ||
    cleanProfilePos === "businessanalyst" ||
    cleanProfilePos === "analyst" ||
    profileLevelClean === "admin";
    
  if (isBusinessAnalyst) {
    if (profileLevelClean !== "admin") {
      profile.position = "Business Analyst";
    }
    var asmSubordinates = [];
    if (profileLevelClean !== "admin") {
      employees.forEach(function(emp) {
        if (emp.name !== "" && emp.position === "Area Sales Manager") {
          if (!asmSubordinates.includes(emp.name)) {
            asmSubordinates.push(emp.name);
          }
        }
      });
    }
    profile.subordinates = asmSubordinates;
    return { status: "success", data: profile };
  }

  var subs = [];
  var queue = Array.from(userAliases);
  var visited = new Set(queue);

  while (queue.length > 0) {
    var currentUpline = queue.shift();
    employees.forEach(function(emp) {
      var empNameRaw = String(emp.name || "").trim();
      var empNameLower = empNameRaw.toLowerCase();
      var empEmailRaw = String(emp.email || "").trim();
      var empEmailLower = empEmailRaw.toLowerCase();
      var empUplineRaw = String(emp.upline || "").trim();
      var empUplineLower = empUplineRaw.toLowerCase();

      if (empUplineLower !== "") {
        var isMatch =
          empUplineLower === currentUpline ||
          empUplineLower.includes(currentUpline) ||
          currentUpline.includes(empUplineLower);
        if (isMatch) {
          var addedAny = false;
          if (empNameLower !== "" && !visited.has(empNameLower)) {
            visited.add(empNameLower);
            queue.push(empNameLower);
            addedAny = true;
          }
          if (empEmailLower !== "" && !visited.has(empEmailLower)) {
            visited.add(empEmailLower);
            queue.push(empEmailLower);
            addedAny = true;
          }
          if (addedAny) {
            var displayName = empNameRaw !== "" ? empNameRaw : empEmailRaw;
            if (displayName !== "" && !subs.includes(displayName)) {
              subs.push(displayName);
            }
          }
        }
      }
    });
  }

  profile.subordinates = subs;
  return { status: "success", data: profile };
}

/**
 * Handler: Ambil aturan akses
 */
function handleGetAccessRules() {
  try {
    var empData = getSheetValues("employee");
    var uniquePositions = [];
    if (empData && empData.length > 1) {
      var headers = empData[0];
      var posIdx = headers.findIndex(function(h) { return /position|jabatan/i.test(String(h).trim()); });
      var nameIdx = headers.findIndex(function(h) { return /nama|name|pic/i.test(String(h).trim()); });
      for (var i = 1; i < empData.length; i++) {
        var row = empData[i];
        if (nameIdx !== -1 && !row[nameIdx]) continue;
        var rawPos = posIdx !== -1 ? row[posIdx] : "";
        var normalized = normalizePosition(rawPos);
        if (normalized && !uniquePositions.includes(normalized)) {
          uniquePositions.push(normalized);
        }
      }
    }
    
    var defaults = ["Business Analyst", "Vegetables Sales Manager", "Area Sales Manager", "Sales Agronomist", "Business Solution"];
    defaults.forEach(function(d) {
      if (!uniquePositions.includes(d)) {
        uniquePositions.push(d);
      }
    });

    var data = getSheetValues("access");
    var rules = {};
    var existingPositions = [];

    if (data && data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var position = String(row[0]).trim();
        if (!position) continue;
        existingPositions.push(position);
        rules[position] = {
          home: row[1] === true || String(row[1]).toUpperCase() === "TRUE",
          partner: row[2] === true || String(row[2]).toUpperCase() === "TRUE",
          stock: row[3] === true || String(row[3]).toUpperCase() === "TRUE",
          pog: row[4] === true || String(row[4]).toUpperCase() === "TRUE",
          overview: row[5] === true || String(row[5]).toUpperCase() === "TRUE",
          temp: row[6] === true || String(row[6]).toUpperCase() === "TRUE",
          access: row[7] === true || String(row[7]).toUpperCase() === "TRUE"
        };
      }
    }

    var missingPositions = uniquePositions.filter(function(p) { return !existingPositions.includes(p); });
    if (missingPositions.length > 0) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("access");
      for (var k = 0; k < missingPositions.length; k++) {
        var p = missingPositions[k];
        var home = "TRUE";
        var partner = "TRUE";
        var stock = "TRUE";
        var pog = "TRUE";
        var overview = "FALSE";
        var temp = "FALSE";
        var access = "FALSE";

        if (p === "Business Analyst") {
          overview = "TRUE";
          temp = "TRUE";
          access = "TRUE";
        }

        if (sheet) {
          sheet.appendRow([p, home, partner, stock, pog, overview, temp, access]);
        }

        rules[p] = {
          home: home === "TRUE",
          partner: partner === "TRUE",
          stock: stock === "TRUE",
          pog: pog === "TRUE",
          overview: overview === "TRUE",
          temp: temp === "TRUE",
          access: access === "TRUE"
        };
      }
    }

    return { status: "success", data: rules };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

/**
 * Handler: Ambil semua data awal dalam satu request parallel (mocked/sequential di GAS)
 */
function handleGetInitialData(user) {
  try {
    var profileJson = handleGetUserProfile(user);
    var employeesJson = handleGetEmployees();
    var channelsJson = handleGetChannels(user);
    var workingDataJson = handleGetWorkingData(user);
    var drSalesDataJson = handleGetDrSalesData(user);
    var accessRulesJson = handleGetAccessRules();

    return {
      status: "success",
      data: {
        profile: profileJson.status === "success" ? profileJson.data : null,
        employees: employeesJson.status === "success" ? employeesJson.data : [],
        channels: channelsJson.status === "success" ? channelsJson.data : [],
        workingData: workingDataJson.status === "success" ? workingDataJson.data : [],
        drSalesData: drSalesDataJson.status === "success" ? drSalesDataJson.data : [],
        accessRules: accessRulesJson.status === "success" ? accessRulesJson.data : {}
      }
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Helper: Normalisasi Jabatan/Posisi
 */
function normalizePosition(pos) {
  if (!pos) return "Unknown";
  var clean = cleanForMatch(pos);
  if (clean === "businessanalyst" || clean === "analyst") return "Business Analyst";
  if (clean === "areasalesmanager" || clean === "asm") return "Area Sales Manager";
  if (clean === "vegetablessalesmanager" || clean === "vsm") return "Vegetables Sales Manager";
  if (clean === "salesmanager" || clean === "sm") return "Sales Manager";
  if (clean === "salesagronomist" || clean === "sa") return "Sales Agronomist";
  if (clean === "businesssolution" || clean === "bs") return "Business Solution";
  if (clean === "countryhead") return "Country Head";
  if (clean === "commerciallead") return "Commercial Lead";

  if (clean.includes("businessanalyst")) return "Business Analyst";
  if (clean.includes("areasalesmanager") || clean.includes("asm")) return "Area Sales Manager";
  if (clean.includes("vegetablessalesmanager")) return "Vegetables Sales Manager";
  return String(pos).trim();
}

/**
 * Helper: Formatter Tanggal DD/MMM/YY
 */
function formatMyDate(dateObj) {
  if (!dateObj || dateObj === "" || dateObj === "N/A" || dateObj === "-" || dateObj === "null" || dateObj === "undefined") {
    return "N/A";
  }
  var date = parseGasDate(dateObj);
  if (date && !isNaN(date.getTime()) && date.getTime() !== 0) {
    var day = String(date.getDate()).padStart(2, "0");
    var months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    var year = String(date.getFullYear()).substring(2);
    return day + "/" + months[date.getMonth()] + "/" + year;
  }
  var str = String(dateObj).trim();
  return str || "N/A";
}

/**
 * Helper: Pencarian karyawan
 */
function findEmployeeRow(user, empData) {
  if (!user || !empData || empData.length <= 1) return null;
  var headers = empData[0];
  var emailIdx = headers.findIndex(function(h) { return /email/i.test(String(h).trim()); });
  var userIdx = headers.findIndex(function(h) { return /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()); });
  var nameIdx = headers.findIndex(function(h) { return /nama|name|pic/i.test(String(h).trim()); });

  var targetClean = cleanForMatch(user);

  for (var i = 1; i < empData.length; i++) {
    var row = empData[i];
    var rowName = nameIdx !== -1 ? cleanForMatch(row[nameIdx]) : "";
    var rowEmail = emailIdx !== -1 ? cleanForMatch(row[emailIdx]) : "";
    var rowUser = userIdx !== -1 ? cleanForMatch(row[userIdx]) : "";

    if (rowUser === targetClean || rowName === targetClean || rowEmail === targetClean) {
      return row;
    }
  }
  return null;
}

/**
 * Overwrite a sheet's content (clear and write values)
 */
function updateSheetValues(sheetName, values) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  sheet.clearContents();
  if (values && values.length > 0 && values[0].length > 0) {
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }
  return true;
}

/**
 * Get active Jakarta date
 */
function getJakartaDate() {
  var formatted = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss");
  return new Date(formatted);
}

/**
 * Parse date string/object safely in Apps Script
 */
function parseGasDate(val) {
  if (!val) return new Date(0);
  if (val instanceof Date) return val;
  var str = String(val).trim();
  var d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  if (str.indexOf("/") !== -1) {
    var parts = str.split(/[\s/:]+/);
    if (parts.length >= 3) {
      var dPart = parseInt(parts[0], 10);
      var mPart = parts[1];
      var yPart = parseInt(parts[2], 10);
      if (yPart < 100) yPart += 2000;
      var m = parseInt(mPart, 10) - 1;
      if (isNaN(m)) {
        var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
        var indoMonths = ["jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des"];
        var lowerM = mPart.toLowerCase();
        m = months.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
        if (m === -1) {
          m = indoMonths.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
        }
        if (m === -1) m = 0;
      }
      return new Date(yPart, m, dPart);
    }
  }
  if (str.indexOf("-") !== -1) {
    var parts = str.split(/[\s\-:]+/);
    if (parts.length >= 3) {
      var dPart = parseInt(parts[0], 10);
      var mPart = parts[1];
      var yPart = parseInt(parts[2], 10);
      if (yPart < 100) yPart += 2000;
      var m = parseInt(mPart, 10) - 1;
      if (isNaN(m)) {
        var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
        var indoMonths = ["jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des"];
        var lowerM = mPart.toLowerCase();
        m = months.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
        if (m === -1) {
          m = indoMonths.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
        }
        if (m === -1) m = 0;
      }
      return new Date(yPart, m, dPart);
    }
  }
  return new Date(0);
}

/**
 * Get month index from timestamp
 */
function getMonthIndexFromDateString(dateStr) {
  if (!dateStr) return new Date().getMonth();
  if (dateStr instanceof Date) return dateStr.getMonth();
  var str = String(dateStr).trim();
  if (str.indexOf("/") !== -1) {
    var parts = str.split(/[\s/:]+/);
    if (parts.length >= 2) {
      var mVal = parseInt(parts[1], 10);
      if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) return mVal - 1;
      var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      var lowerM = parts[1].toLowerCase();
      var m = months.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
      if (m !== -1) return m;
    }
  }
  if (str.indexOf("-") !== -1) {
    var parts = str.split(/[\s\-:]+/);
    if (parts.length >= 2) {
      var mVal = parseInt(parts[1], 10);
      if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) return mVal - 1;
      var months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      var lowerM = parts[1].toLowerCase();
      var m = months.findIndex(function(name) { return lowerM.indexOf(name) === 0; });
      if (m !== -1) return m;
    }
  }
  return new Date().getMonth();
}

/**
 * Safe bounds helper
 */
function colIndexInBounds(idx, row) {
  return idx !== -1 && idx < row.length;
}

/**
 * Resolve employee province/area
 */
function getUserProvince(userName, empData) {
  if (!userName || !empData || empData.length <= 1) return "";
  var headers = empData[0];
  var provIdx = headers.findIndex(function(h) {
    return /province|provinsi/i.test(String(h).trim());
  });
  var areaIdx = headers.findIndex(function(h) {
    return /area/i.test(String(h).trim());
  });
  var row = findEmployeeRow(userName, empData);
  if (row) {
    if (provIdx !== -1 && row[provIdx] !== "" && row[provIdx] !== undefined) {
      return String(row[provIdx]).trim();
    }
    if (areaIdx !== -1 && row[areaIdx] !== "" && row[areaIdx] !== undefined) {
      return String(row[areaIdx]).trim();
    }
  }
  return "";
}

/**
 * Extract expiration date from Lot number
 */
function extractExpFromLot(lotNo) {
  if (!lotNo) return "N/A";
  var raw = String(lotNo).trim().toUpperCase();
  if (!raw || raw === "N/A" || raw === "-") return "N/A";

  var months = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
    "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
  ];

  // Clean to alphanumeric
  var clean = raw.replace(/[^A-Z0-9]/g, "");

  // 1. Check for 8 digits DDMMYYYY
  var digit8Matches = clean.match(/\d{8}/g);
  if (digit8Matches) {
    for (var i = 0; i < digit8Matches.length; i++) {
      var mStr = digit8Matches[i];
      var d = parseInt(mStr.slice(0, 2), 10);
      var m = parseInt(mStr.slice(2, 4), 10);
      var y = parseInt(mStr.slice(4, 8), 10);

      if (m > 12 && d >= 1 && d <= 12) {
        var temp = d; d = m; m = temp;
      }
      if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 2020 && y <= 2050) {
        var dayStr = String(d).padStart(2, "0");
        var yearStr = String(y).slice(-2);
        return dayStr + "/" + months[m - 1] + "/" + yearStr;
      }
    }
  }

  // 2. Check for 6 digits (DDMMYY or YYMMDD)
  var digit6Matches = clean.match(/\d{6}/g);
  if (digit6Matches) {
    for (var i = 0; i < digit6Matches.length; i++) {
      var mStr = digit6Matches[i];
      var d = parseInt(mStr.slice(0, 2), 10);
      var m = parseInt(mStr.slice(2, 4), 10);
      var y = parseInt(mStr.slice(4, 6), 10);

      if (y >= 20 && y <= 50) {
        if (m > 12 && d >= 1 && d <= 12) {
          var temp = d; d = m; m = temp;
        }
        if (d >= 1 && d <= 31 && m >= 1 && m <= 12) {
          var dayStr = String(d).padStart(2, "0");
          var yearStr = String(y).padStart(2, "0");
          return dayStr + "/" + months[m - 1] + "/" + yearStr;
        }
      }

      var yy = parseInt(mStr.slice(0, 2), 10);
      var mm = parseInt(mStr.slice(2, 4), 10);
      var dd = parseInt(mStr.slice(4, 6), 10);
      if (yy >= 20 && yy <= 50 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
        var dayStr = String(dd).padStart(2, "0");
        var yearStr = String(yy).padStart(2, "0");
        return dayStr + "/" + months[mm - 1] + "/" + yearStr;
      }
    }
  }

  // 3. Fallback: match 6 digits in raw string
  var match = raw.match(/(\d{2})(\d{2})(\d{2})/);
  if (match) {
    var d = parseInt(match[1], 10);
    var m = parseInt(match[2], 10);
    var y = parseInt(match[3], 10);

    if (m > 12 && d >= 1 && d <= 12) {
      var temp = d; d = m; m = temp;
    }

    if (d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 20 && y <= 50) {
      var dayStr = String(d).padStart(2, "0");
      var yearStr = String(y).padStart(2, "0");
      return dayStr + "/" + months[m - 1] + "/" + yearStr;
    }
  }

  return "N/A";
}

/**
 * Batch activity processor mirroring Express handler
 */
function handleBatchActivity(body) {
  try {
    var data = getSheetValues("working");
    if (!data) return { status: "error", message: "Sheet 'working' tidak ditemukan" };

    var headers = data[0];

    // Ensure we have Tgl and Category columns in the sheet
    var hasTgl = false;
    var hasCat = false;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === "tgl" || h === "tanggal") hasTgl = true;
      if (h === "category" || h === "kategori") hasCat = true;
    }

    if (!hasTgl) {
      headers.push("Tgl");
    }
    if (!hasCat) {
      headers.push("Category");
    }

    var idxPog = headers.findIndex(function(h) {
      return /^pog$|^selisih$/i.test(String(h).trim());
    });
    if (idxPog === -1) {
      headers.push("POG");
      idxPog = headers.length - 1;
    }

    // Always make sure each row in data is padded to headers.length
    for (var i = 1; i < data.length; i++) {
      while (data[i].length < headers.length) {
        data[i].push("");
      }
    }

    var getIdx = function(patterns) {
      return headers.findIndex(function(h) { return patterns.test(String(h).trim()); });
    };

    var getIdxList = function(patterns) {
      var indices = [];
      for (var i = 0; i < headers.length; i++) {
        if (patterns.test(String(headers[i]).trim())) {
          indices.push(i);
        }
      }
      return indices;
    };

    var timeCols = getIdxList(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i);
    var catCols = getIdxList(/^category$|^kategori$|^klasifikasi$|^cat$/i);

    var idx = {
      time: getIdx(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i),
      kiosk: getIdx(/^channel$|^kiosk$/i),
      cat: getIdx(/^category$|^kategori$|^klasifikasi$|^cat$/i),
      user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
      lot: getIdx(/^lot package$|^lot$/i),
      qty: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
      area: getIdx(/^area$|^region$/i),
      desc: getIdx(/^hybrid$|^material$/i),
      exp: getIdx(/^exp date$|^expired$/i),
      agingMonth: getIdx(/^aging \(month\)/i),
      cond: getIdx(/^condition$|^kondisi$/i),
      crops: getIdx(/^crops$/i),
      dr: getIdx(/^shipping date$|^dr date$/i),
      agingExp: getIdx(/^aging to exp$/i),
      cluster: getIdx(/^cluster$/i),
      pog: idxPog
    };

    var now = getJakartaDate();
    var pad = function(n) { return String(n).padStart(2, "0"); };
    var timestamp = pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

    // Backfill existing empty timestamp and category data as requested
    var channelData = getSheetValues("channel") || [];
    var channelCategoryMap = {};
    if (channelData && channelData.length > 1) {
      var chanHeaders = channelData[0];
      var chanIdxName = chanHeaders.findIndex(function(h) { return /name|nama|channel|kiosk/i.test(String(h).trim()); });
      var chanIdxCat = chanHeaders.findIndex(function(h) { return /kategori|category|klasifikasi/i.test(String(h).trim()); });
      if (chanIdxName !== -1 && chanIdxCat !== -1) {
        for (var c = 1; c < channelData.length; c++) {
          var cRow = channelData[c];
          var cName = String(cRow[chanIdxName] || "").trim().toLowerCase();
          var cCat = String(cRow[chanIdxCat] || "").trim();
          if (cName) {
            channelCategoryMap[cName] = cCat;
          }
        }
      }
    }

    function getGasCategory(kioskName, itemCat) {
      if (itemCat && itemCat !== "-" && itemCat.toLowerCase() !== "uncategorized" && itemCat.toLowerCase() !== "n/a") {
        return itemCat;
      }
      var cleanKName = String(kioskName || "").trim().toLowerCase();
      if (channelCategoryMap[cleanKName]) {
        return channelCategoryMap[cleanKName];
      }
      if (cleanKName.indexOf("distributor") !== -1 || cleanKName.indexOf("distr") !== -1 || cleanKName.indexOf("gudang") !== -1 || cleanKName.indexOf("warehouse") !== -1 || cleanKName.indexOf("subd") !== -1) {
        return "Distributor";
      } else if (/\br1\b|toko r1|mitra r1|kiosk r1/i.test(cleanKName)) {
        return "R1";
      } else if (/\br2\b|toko r2|mitra r2|kiosk r2/i.test(cleanKName)) {
        return "R2";
      }
      return "R1";
    }

    var kioskIdxForBackfill = idx.kiosk;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 1. Fill empty timestamp/tgl columns
      var hasValidTime = false;
      timeCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
          hasValidTime = true;
        }
      });
      if (!hasValidTime) {
        timeCols.forEach(function(colIdx) {
          row[colIdx] = timestamp;
        });
      } else {
        var existingTimeVal = "";
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
            existingTimeVal = String(row[colIdx]).trim();
          }
        });
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] === undefined || row[colIdx] === null || String(row[colIdx]).trim() === "") {
            row[colIdx] = existingTimeVal;
          }
        });
      }

      // 2. Fill empty category columns
      var kioskNameVal = kioskIdxForBackfill !== -1 ? String(row[kioskIdxForBackfill] || "").trim() : "";
      var existingCatVal = "";
      catCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "" && String(row[colIdx]).trim().toLowerCase() !== "uncategorized") {
          existingCatVal = String(row[colIdx]).trim();
        }
      });

      if (!existingCatVal && kioskNameVal) {
        existingCatVal = getGasCategory(kioskNameVal);
      }
      if (!existingCatVal) {
        existingCatVal = "R1";
      }

      catCols.forEach(function(colIdx) {
        row[colIdx] = existingCatVal;
      });
    }

    var monthIndices = getMonthIndices(headers);
    var updMonthIndices = getUpdMonthIndices(headers);

    var currentMonthRowsMap = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var k = idx.kiosk !== -1 ? cleanForMatch(row[idx.kiosk]) : "";
      var l = idx.lot !== -1 ? cleanForMatch(row[idx.lot]) : "";
      var h = idx.desc !== -1 ? cleanForMatch(row[idx.desc]) : "";
      var u = idx.user !== -1 ? cleanForMatch(row[idx.user]) : "";
      var key = k + "_" + l + "_" + h + "_" + u;
      if (k && l) {
        var rowDate = new Date(0);
        if (idx.time !== -1 && row[idx.time]) {
          rowDate = parseGasDate(row[idx.time]);
        }
        if (!currentMonthRowsMap[key] || rowDate.getTime() > currentMonthRowsMap[key].date.getTime()) {
          currentMonthRowsMap[key] = { index: i + 1, date: rowDate };
        }
      }
    }

    var now = getJakartaDate();
    var pad = function(n) { return String(n).padStart(2, "0"); };
    var timestamp = pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

    var empData = getSheetValues("employee") || [];

    if (body.items && body.items.length > 0) {
      for (var j = 0; j < body.items.length; j++) {
        var item = body.items[j];
        var k = cleanForMatch(item.kiosk);
        var l = cleanForMatch(item.lot);
        var h = cleanForMatch(item.hybrid);
        var u = cleanForMatch(item.user || body.user);
        var key = k + "_" + l + "_" + h + "_" + u;
        var existingMatch = currentMonthRowsMap[key];
        var existingRow = existingMatch ? existingMatch.index : null;
        var agingExpVal = "";
        if (item.expired && item.expired !== "N/A") {
          var expD = new Date(item.expired);
          if (!isNaN(expD.getTime())) {
            var today = getJakartaDate();
            today.setHours(0, 0, 0, 0);
            agingExpVal = String(Math.round((expD.getTime() - today.getTime()) / (1000 * 3600 * 24) / 30.416));
          }
        }
        var clusterVal = "";
        if (item.aging && item.aging !== "-" && !isNaN(Number(item.aging))) {
          var aVal = Number(item.aging);
          if (aVal <= 2) clusterVal = "0-2";
          else if (aVal <= 4) clusterVal = "2-4";
          else if (aVal <= 6) clusterVal = "4-6";
          else if (aVal <= 9) clusterVal = "6-9";
          else if (aVal <= 12) clusterVal = "9-12";
          else clusterVal = ">12";
        }
        var itemMonthIdx = getMonthIndexFromDateString(timestamp);
        var monthColIdx = monthIndices[itemMonthIdx];

        var prevMonthStock = 0;
        if (existingRow) {
          var prevMonthIdx = (itemMonthIdx - 1 + 12) % 12;
          var prevMonthColIdx = monthIndices[prevMonthIdx];
          if (prevMonthColIdx !== -1) {
            prevMonthStock = Number(data[existingRow - 1][prevMonthColIdx]) || 0;
          } else {
            prevMonthStock = Number(item.originalStock) || 0;
          }
        } else {
          prevMonthStock = Number(item.originalStock) || 0;
        }

        if (item.condition === "habis" || Number(item.stock) === 0) {
          item.stock = 0;
        }
        var pogVal = prevMonthStock - (Number(item.stock) || 0);

        if (existingRow) {
          var rowIndex = existingRow - 1;
          if (monthColIdx !== -1) {
            data[rowIndex][monthColIdx] = item.stock;
          }
          var updMonthColIdx = updMonthIndices[itemMonthIdx];
          if (updMonthColIdx !== -1) {
            data[rowIndex][updMonthColIdx] = "sales";
          }

          if (idx.qty !== -1) data[rowIndex][idx.qty] = item.stock;
          if (idx.cond !== -1) data[rowIndex][idx.cond] = item.condition;
          
          timeCols.forEach(function(colIdx) {
            data[rowIndex][colIdx] = timestamp;
          });

          if (idx.user !== -1) data[rowIndex][idx.user] = item.user || body.user;
          var resolvedArea = getUserProvince(item.user || body.user, empData) || body.area || "";
          if (idx.area !== -1 && resolvedArea) data[rowIndex][idx.area] = resolvedArea;
          if (idx.agingExp !== -1 && agingExpVal !== "") data[rowIndex][idx.agingExp] = agingExpVal;
          if (idx.cluster !== -1 && clusterVal !== "") data[rowIndex][idx.cluster] = clusterVal;
          
          var rowCatVal = getGasCategory(item.kiosk, item.category || body.category);
          catCols.forEach(function(colIdx) {
            data[rowIndex][colIdx] = rowCatVal;
          });

          if (idx.pog !== -1) data[rowIndex][idx.pog] = pogVal;
        } else {
          var newRow = [];
          for (var c = 0; c < headers.length; c++) newRow.push("");
          
          timeCols.forEach(function(colIdx) {
            newRow[colIdx] = timestamp;
          });

          if (idx.kiosk !== -1) newRow[idx.kiosk] = item.kiosk;
          
          var rowCatVal = getGasCategory(item.kiosk, item.category || body.category);
          catCols.forEach(function(colIdx) {
            newRow[colIdx] = rowCatVal;
          });

          if (idx.user !== -1) newRow[idx.user] = item.user || body.user;
          if (idx.lot !== -1) newRow[idx.lot] = String(item.lot).toUpperCase();

          if (monthColIdx !== -1) {
            newRow[monthColIdx] = item.stock;
          }
          var updMonthColIdx = updMonthIndices[itemMonthIdx];
          if (updMonthColIdx !== -1) {
            newRow[updMonthColIdx] = "sales";
          }

          if (idx.qty !== -1) newRow[idx.qty] = item.stock;
          var resolvedArea = getUserProvince(item.user || body.user, empData) || body.area || "";
          if (idx.area !== -1) newRow[idx.area] = resolvedArea;
          if (idx.desc !== -1) newRow[idx.desc] = item.hybrid;
          if (idx.crops !== -1) newRow[idx.crops] = item.crops || "";
          if (idx.dr !== -1) newRow[idx.dr] = item.drDate || "";
          if (idx.exp !== -1) newRow[idx.exp] = item.expired;
          if (idx.agingMonth !== -1) newRow[idx.agingMonth] = item.aging;
          if (idx.cond !== -1) newRow[idx.cond] = item.condition;
          if (idx.agingExp !== -1) newRow[idx.agingExp] = agingExpVal;
          if (idx.cluster !== -1) newRow[idx.cluster] = clusterVal;
          if (idx.pog !== -1) newRow[idx.pog] = pogVal;

          for (var m = 0; m < 12; m++) {
            var colIdx = monthIndices[m];
            if (colIdx !== -1 && colIdx !== monthColIdx) {
              newRow[colIdx] = 0;
            }
          }
          data.push(newRow);
        }
      }
      updateSheetValues("working", data);
    }
    return { status: "success" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

/**
 * Consolidate working sheet data mirroring Express handler
 */
function handleConsolidateDatabase(body) {
  try {
    var data = getSheetValues("working");
    var drData = getSheetValues("dr");
    var empDataRaw = getSheetValues("employee");

    if (!data || data.length <= 1) {
      return { status: "success", message: "Tidak ada data untuk dikonsolidasi" };
    }

    var empData = empDataRaw || [];
    var headers = data[0];

    // Ensure we have Tgl and Category columns in the sheet
    var hasTgl = false;
    var hasCat = false;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === "tgl" || h === "tanggal") hasTgl = true;
      if (h === "category" || h === "kategori") hasCat = true;
    }

    if (!hasTgl) {
      headers.push("Tgl");
    }
    if (!hasCat) {
      headers.push("Category");
    }

    var idxPog = headers.findIndex(function(h) {
      return /^pog$|^selisih$/i.test(String(h).trim());
    });
    if (idxPog === -1) {
      headers.push("POG");
      idxPog = headers.length - 1;
    }

    // Always make sure each row in data is padded to headers.length
    for (var i = 1; i < data.length; i++) {
      while (data[i].length < headers.length) {
        data[i].push("");
      }
    }

    var getIdx = function(patterns) {
      return headers.findIndex(function(h) { return patterns.test(String(h).trim()); });
    };

    var getIdxList = function(patterns) {
      var indices = [];
      for (var i = 0; i < headers.length; i++) {
        if (patterns.test(String(headers[i]).trim())) {
          indices.push(i);
        }
      }
      return indices;
    };

    var timeCols = getIdxList(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i);
    var catCols = getIdxList(/^category$|^kategori$|^klasifikasi$|^cat$/i);

    var idx = {
      time: getIdx(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i),
      kiosk: getIdx(/^channel$|^kiosk$/i),
      cat: getIdx(/^category$|^kategori$|^klasifikasi$|^cat$/i),
      user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
      lot: getIdx(/^lot package$|^lot$/i),
      qty: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
      area: getIdx(/^area$|^region$/i),
      desc: getIdx(/^hybrid$|^material$/i),
      exp: getIdx(/^exp date$|^expired$/i),
      agingMonth: getIdx(/^aging \(month\)/i),
      cond: getIdx(/^condition$|^kondisi$/i),
      crops: getIdx(/^crops$/i),
      dr: getIdx(/^shipping date$|^dr date$/i),
      agingExp: getIdx(/^aging to exp$/i),
      cluster: getIdx(/^cluster$/i),
      pog: idxPog
    };

    var now = getJakartaDate();
    var pad = function(n) { return String(n).padStart(2, "0"); };
    var timestamp = pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

    // Fetch channel category map for backfilling categories
    var channelData = getSheetValues("channel") || [];
    var channelCategoryMap = {};
    if (channelData && channelData.length > 1) {
      var chanHeaders = channelData[0];
      var chanIdxName = chanHeaders.findIndex(function(h) { return /name|nama|channel|kiosk/i.test(String(h).trim()); });
      var chanIdxCat = chanHeaders.findIndex(function(h) { return /kategori|category|klasifikasi/i.test(String(h).trim()); });
      if (chanIdxName !== -1 && chanIdxCat !== -1) {
        for (var c = 1; c < channelData.length; c++) {
          var cRow = channelData[c];
          var cName = String(cRow[chanIdxName] || "").trim().toLowerCase();
          var cCat = String(cRow[chanIdxCat] || "").trim();
          if (cName) {
            channelCategoryMap[cName] = cCat;
          }
        }
      }
    }

    function getGasCategory(kioskName, itemCat) {
      if (itemCat && itemCat !== "-" && itemCat.toLowerCase() !== "uncategorized" && itemCat.toLowerCase() !== "n/a") {
        return itemCat;
      }
      var cleanKName = String(kioskName || "").trim().toLowerCase();
      if (channelCategoryMap[cleanKName]) {
        return channelCategoryMap[cleanKName];
      }
      if (cleanKName.indexOf("distributor") !== -1 || cleanKName.indexOf("distr") !== -1 || cleanKName.indexOf("gudang") !== -1 || cleanKName.indexOf("warehouse") !== -1 || cleanKName.indexOf("subd") !== -1) {
        return "Distributor";
      } else if (/\br1\b|toko r1|mitra r1|kiosk r1/i.test(cleanKName)) {
        return "R1";
      } else if (/\br2\b|toko r2|mitra r2|kiosk r2/i.test(cleanKName)) {
        return "R2";
      }
      return "R1";
    }

    var kioskIdxForBackfill = idx.kiosk;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 1. Fill empty timestamp/tgl columns
      var hasValidTime = false;
      timeCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
          hasValidTime = true;
        }
      });
      if (!hasValidTime) {
        timeCols.forEach(function(colIdx) {
          row[colIdx] = timestamp;
        });
      } else {
        var existingTimeVal = "";
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
            existingTimeVal = String(row[colIdx]).trim();
          }
        });
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] === undefined || row[colIdx] === null || String(row[colIdx]).trim() === "") {
            row[colIdx] = existingTimeVal;
          }
        });
      }

      // 2. Fill empty category columns
      var kioskNameVal = kioskIdxForBackfill !== -1 ? String(row[kioskIdxForBackfill] || "").trim() : "";
      var existingCatVal = "";
      catCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "" && String(row[colIdx]).trim().toLowerCase() !== "uncategorized") {
          existingCatVal = String(row[colIdx]).trim();
        }
      });

      if (!existingCatVal && kioskNameVal) {
        existingCatVal = getGasCategory(kioskNameVal);
      }
      if (!existingCatVal) {
        existingCatVal = "R1";
      }

      catCols.forEach(function(colIdx) {
        row[colIdx] = existingCatVal;
      });
    }

    var monthIndices = getMonthIndices(headers);
    var updMonthIndices = getUpdMonthIndices(headers);

    var isValidVal = function(v) {
      return (
        v !== undefined &&
        v !== null &&
        String(v).trim() !== "" &&
        String(v).trim().toUpperCase() !== "N/A" &&
        String(v).trim() !== "-"
      );
    };

    var lotLookup = {};
    if (drData && drData.length > 1) {
      var drHeaders = drData[0];
      var drIdx = {
        lot: drHeaders.findIndex(function(h) { return /lot/i.test(String(h).trim()); }),
        dr: drHeaders.findIndex(function(h) { return /dr date|shipping date/i.test(String(h).trim()); }),
        exp: drHeaders.findIndex(function(h) { return /exp date|expired/i.test(String(h).trim()); })
      };
      if (drIdx.lot !== -1) {
        for (var j = 1; j < drData.length; j++) {
          var drRow = drData[j];
          var lNo = String(drRow[drIdx.lot] || "").trim().toUpperCase();
          if (lNo && !lotLookup[lNo]) {
            lotLookup[lNo] = {
              drDate: (drIdx.dr !== -1 && drRow[drIdx.dr]) ? formatMyDate(drRow[drIdx.dr]) : "",
              expDate: (drIdx.exp !== -1 && drRow[drIdx.exp]) ? formatMyDate(drRow[drIdx.exp]) : ""
            };
          }
        }
      }
    }

    var grouped = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && idx.kiosk !== -1 && !row[idx.kiosk]) continue;

      var kioskVal = (idx.kiosk !== -1) ? String(row[idx.kiosk] || "").trim() : "";
      var lotVal = (idx.lot !== -1) ? String(row[idx.lot] || "").trim() : "";
      var descVal = (idx.desc !== -1) ? String(row[idx.desc] || "").trim() : "";
      var userVal = (idx.user !== -1) ? String(row[idx.user] || "").trim() : "";

      var groupKey = cleanForMatch(kioskVal) + "_" + cleanForMatch(lotVal) + "_" + cleanForMatch(descVal) + "_" + cleanForMatch(userVal);
      var timestampStr = idx.time !== -1 ? row[idx.time] : "";
      var rowDate = timestampStr ? parseGasDate(timestampStr) : new Date(0);
      var rowMonthIdx = getMonthIndexFromDateString(timestampStr);

      var monthValsSrc = [];
      for (var m = 0; m < 12; m++) monthValsSrc.push(0);
      monthIndices.forEach(function(colIdx, mIdx) {
        if (colIdx !== -1 && colIdx < row.length && row[colIdx] !== "") {
          monthValsSrc[mIdx] = Number(row[colIdx]) || 0;
        }
      });

      var updValsSrc = [];
      for (var m = 0; m < 12; m++) updValsSrc.push("");
      updMonthIndices.forEach(function(colIdx, mIdx) {
        if (
          colIdx !== -1 &&
          colIdx < row.length &&
          row[colIdx] !== undefined &&
          row[colIdx] !== null &&
          row[colIdx] !== ""
        ) {
          updValsSrc[mIdx] = String(row[colIdx]).trim();
        }
      });

      var totalMonthVals = 0;
      for (var m = 0; m < 12; m++) totalMonthVals += monthValsSrc[m];
      var qtyVal = idx.qty !== -1 ? (Number(row[idx.qty]) || 0) : 0;
      if (totalMonthVals === 0 && qtyVal > 0) {
        monthValsSrc[rowMonthIdx] = qtyVal;
      }

      var lotUpper = String(lotVal).trim().toUpperCase();
      var currentExp = "";
      var currentDr = "";

      // 1. Prioritize lookups from lotLookup (the source of truth 'dr' sheet)
      if (lotLookup[lotUpper]) {
        if (isValidVal(lotLookup[lotUpper].drDate)) {
          currentDr = lotLookup[lotUpper].drDate;
        }
        if (isValidVal(lotLookup[lotUpper].expDate)) {
          currentExp = lotLookup[lotUpper].expDate;
        }
      }

      // 2. Fallback to existing row values if not available in lotLookup
      if (!isValidVal(currentExp)) {
        currentExp = (idx.exp !== -1 && colIndexInBounds(idx.exp, row)) ? row[idx.exp] : "";
      }
      if (!isValidVal(currentDr)) {
        currentDr = (idx.dr !== -1 && colIndexInBounds(idx.dr, row)) ? row[idx.dr] : "";
      }

      // 3. Last fallback (extract exp from lot and default dr)
      if (!isValidVal(currentExp)) {
        currentExp = extractExpFromLot(lotUpper);
      }
      if (!isValidVal(currentDr)) {
        currentDr = "N/A";
      }

      if (!grouped[groupKey]) {
        grouped[groupKey] = {
          kiosk: kioskVal,
          category: (idx.cat !== -1 && colIndexInBounds(idx.cat, row)) ? String(row[idx.cat] || "").trim() : "",
          lot: lotVal,
          desc: descVal,
          user: userVal,
          timestamp: rowDate,
          originalTimestampStr: timestampStr,
          area: getUserProvince(userVal, empData) || (idx.area !== -1 && colIndexInBounds(idx.area, row) ? row[idx.area] : ""),
          crops: (idx.crops !== -1 && colIndexInBounds(idx.crops, row)) ? row[idx.crops] : "",
          exp: currentExp,
          dr: currentDr,
          agingMonth: (idx.agingMonth !== -1 && colIndexInBounds(idx.agingMonth, row)) ? row[idx.agingMonth] : "",
          cond: (idx.cond !== -1 && colIndexInBounds(idx.cond, row)) ? row[idx.cond] : "tetap",
          agingExp: (idx.agingExp !== -1 && colIndexInBounds(idx.agingExp, row)) ? row[idx.agingExp] : "",
          cluster: (idx.cluster !== -1 && colIndexInBounds(idx.cluster, row)) ? row[idx.cluster] : "",
          pog: (idx.pog !== -1 && colIndexInBounds(idx.pog, row)) ? (Number(row[idx.pog]) || 0) : 0,
          monthlyQty: monthValsSrc,
          updMonthVals: updValsSrc,
          rawIndex: i
        };
      } else {
        var g = grouped[groupKey];
        for (var m = 0; m < 12; m++) {
          g.monthlyQty[m] += monthValsSrc[m];
        }
        if (!g.updMonthVals) {
          g.updMonthVals = [];
          for (var m = 0; m < 12; m++) g.updMonthVals.push("");
        }
        for (var m = 0; m < 12; m++) {
          var v1 = String(g.updMonthVals[m] || "").trim().toLowerCase();
          var v2 = String(updValsSrc[m] || "").trim().toLowerCase();
          if (v1 === "sales" || v2 === "sales") {
            g.updMonthVals[m] = "sales";
          } else if (v1 === "admin" || v2 === "admin") {
            g.updMonthVals[m] = "admin";
          } else {
            g.updMonthVals[m] = "";
          }
        }

        var isNewer = rowDate.getTime() > g.timestamp.getTime();
        if (isNewer) {
          g.timestamp = rowDate;
          g.originalTimestampStr = timestampStr;
        }

        var updateField = function(gKey, rowVal) {
          if (isNewer) {
            if (isValidVal(rowVal)) g[gKey] = rowVal;
          } else {
            if (isValidVal(rowVal) && !isValidVal(g[gKey])) g[gKey] = rowVal;
          }
        };

        if (idx.area !== -1 && colIndexInBounds(idx.area, row)) {
          updateField("area", getUserProvince(userVal, empData) || row[idx.area]);
        }
        if (idx.cat !== -1 && colIndexInBounds(idx.cat, row)) {
          updateField("category", row[idx.cat]);
        }
        if (idx.crops !== -1 && colIndexInBounds(idx.crops, row)) {
          updateField("crops", row[idx.crops]);
        }
        updateField("exp", currentExp);
        updateField("dr", currentDr);
        if (idx.agingMonth !== -1 && colIndexInBounds(idx.agingMonth, row)) {
          updateField("agingMonth", row[idx.agingMonth]);
        }
        if (idx.agingExp !== -1 && colIndexInBounds(idx.agingExp, row)) {
          updateField("agingExp", row[idx.agingExp]);
        }
        if (idx.cluster !== -1 && colIndexInBounds(idx.cluster, row)) {
          updateField("cluster", row[idx.cluster]);
        }

        if (idx.cond !== -1 && colIndexInBounds(idx.cond, row)) {
          var rowVal = row[idx.cond];
          if (isNewer && isValidVal(rowVal) && rowVal !== "tetap") {
            g.cond = rowVal;
          } else if (
            !isNewer &&
            isValidVal(rowVal) &&
            rowVal !== "tetap" &&
            (!g.cond || g.cond === "tetap")
          ) {
            g.cond = rowVal;
          }
        }

        if (idx.pog !== -1 && colIndexInBounds(idx.pog, row)) {
          g.pog += Number(row[idx.pog]) || 0;
        }
      }
    }

    var todayDate = getJakartaDate();
    todayDate.setHours(0, 0, 0, 0);
    var curMonthIdx = todayDate.getMonth();
    var prevMonthIdx = (curMonthIdx - 1 + 12) % 12;

    var newSheetValues = [headers];

    var groupKeys = Object.keys(grouped);
    for (var k = 0; k < groupKeys.length; k++) {
      var g = grouped[groupKeys[k]];
      if (!g.monthlyQty) {
        g.monthlyQty = [];
        for (var m = 0; m < 12; m++) g.monthlyQty.push(0);
      }
      if (!g.updMonthVals) {
        g.updMonthVals = [];
        for (var m = 0; m < 12; m++) g.updMonthVals.push("");
      }
      if (g.updMonthVals[curMonthIdx] !== "sales") {
        g.monthlyQty[curMonthIdx] = g.monthlyQty[prevMonthIdx] || 0;
        g.updMonthVals[curMonthIdx] = "admin";
      }

      var curStock = Number(g.monthlyQty[curMonthIdx]) || 0;
      var prevStock = Number(g.monthlyQty[prevMonthIdx]) || 0;
      g.pog = prevStock - curStock;

      // Force-resolve using the lotLookup source of truth
      var lotUpper = String(g.lot).trim().toUpperCase();
      if (lotLookup[lotUpper]) {
        if (isValidVal(lotLookup[lotUpper].drDate)) {
          g.dr = lotLookup[lotUpper].drDate;
        }
        if (isValidVal(lotLookup[lotUpper].expDate)) {
          g.exp = lotLookup[lotUpper].expDate;
        }
      }
      if (!isValidVal(g.exp)) {
        g.exp = extractExpFromLot(lotUpper);
      }

      if (isValidVal(g.dr) && g.dr !== "N/A") {
        g.dr = formatMyDate(parseGasDate(g.dr));
      }
      if (isValidVal(g.exp) && g.exp !== "N/A") {
        g.exp = formatMyDate(parseGasDate(g.exp));
      }

      if (isValidVal(g.dr) && g.dr !== "N/A") {
        var drD = parseGasDate(g.dr);
        if (drD && !isNaN(drD.getTime()) && drD.getTime() !== 0) {
          var calcAge = Math.round((todayDate.getTime() - drD.getTime()) / (1000 * 3600 * 24) / 30.416);
          g.agingMonth = calcAge >= 0 ? calcAge : 0;
        }
      }

      if (isValidVal(g.exp) && g.exp !== "N/A") {
        var expD = parseGasDate(g.exp);
        if (expD && !isNaN(expD.getTime()) && expD.getTime() !== 0) {
          g.agingExp = Math.round((expD.getTime() - todayDate.getTime()) / (1000 * 3600 * 24) / 30.416);
        }
      }

      if (
        g.agingMonth !== "" &&
        g.agingMonth !== undefined &&
        g.agingMonth !== null &&
        !isNaN(Number(g.agingMonth))
      ) {
        var aVal = Number(g.agingMonth);
        if (aVal <= 2) g.cluster = "0-2";
        else if (aVal <= 4) g.cluster = "2-4";
        else if (aVal <= 6) g.cluster = "4-6";
        else if (aVal <= 9) g.cluster = "6-9";
        else if (aVal <= 12) g.cluster = "9-12";
        else g.cluster = ">12";
      }

      var totalQty = 0;
      for (var m = 0; m < 12; m++) totalQty += g.monthlyQty[m];
      var currentMonthQty = g.monthlyQty[curMonthIdx];
      var prevMonthQty = g.monthlyQty[prevMonthIdx];

      if (totalQty === 0) {
        g.cond = "habis";
      } else {
        if (g.cond === "habis") g.cond = "tetap";
        if (g.cond !== "new" && g.cond !== "baru" && g.cond !== "baru (new)") {
          if (currentMonthQty < prevMonthQty) {
            g.cond = "berkurang";
          } else if (currentMonthQty > prevMonthQty) {
            g.cond = "bertambah";
          } else {
            g.cond = "tetap";
          }
        }
      }

      for (var m = 0; m < 12; m++) {
        var uVal = String(g.updMonthVals[m] || "").trim().toLowerCase();
        if (uVal === "sales") {
          g.updMonthVals[m] = "sales";
        } else if (uVal === "admin") {
          g.updMonthVals[m] = "admin";
        } else {
          g.updMonthVals[m] = "";
        }
      }

      var newRow = [];
      for (var c = 0; c < headers.length; c++) newRow.push("");
      var pad = function(n) { return String(n).padStart(2, "0"); };
      var now = getJakartaDate();
      var timestamp = pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

      var isSales = g.updMonthVals[curMonthIdx] === "sales";
      var resolvedTimestamp = isSales ? (g.originalTimestampStr || timestamp) : timestamp;

      timeCols.forEach(function(colIdx) {
        newRow[colIdx] = resolvedTimestamp;
      });
      if (idx.kiosk !== -1) newRow[idx.kiosk] = g.kiosk;
      
      var rowCatVal = getGasCategory(g.kiosk, g.category);
      catCols.forEach(function(colIdx) {
        newRow[colIdx] = rowCatVal;
      });
      if (idx.user !== -1) newRow[idx.user] = g.user;
      if (idx.lot !== -1) newRow[idx.lot] = String(g.lot).toUpperCase();
      if (idx.qty !== -1) newRow[idx.qty] = g.monthlyQty[curMonthIdx];

      if (idx.area !== -1) newRow[idx.area] = g.area;
      if (idx.desc !== -1) newRow[idx.desc] = g.desc;
      if (idx.crops !== -1) newRow[idx.crops] = g.crops;
      if (idx.dr !== -1) newRow[idx.dr] = g.dr;
      if (idx.exp !== -1) newRow[idx.exp] = g.exp;
      if (idx.agingMonth !== -1) newRow[idx.agingMonth] = g.agingMonth;
      if (idx.cond !== -1) newRow[idx.cond] = g.cond;
      if (idx.agingExp !== -1) newRow[idx.agingExp] = g.agingExp;
      if (idx.cluster !== -1) newRow[idx.cluster] = g.cluster;
      if (idx.pog !== -1) newRow[idx.pog] = g.pog;

      monthIndices.forEach(function(colIdx, mIdx) {
        if (colIdx !== -1) {
          newRow[colIdx] = (g.monthlyQty[mIdx] !== 0 || g.updMonthVals[mIdx] !== "") ? Number(g.monthlyQty[mIdx]) || 0 : "";
        }
      });

      updMonthIndices.forEach(function(colIdx, mIdx) {
        if (colIdx !== -1) {
          newRow[colIdx] = g.updMonthVals[mIdx] || "";
        }
      });

      newSheetValues.push(newRow);
    }

    updateSheetValues("working", newSheetValues);
    return { status: "success", message: "Konsolidasi berhasil dilakukan" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Handle Employee update/insert
 */
function handleUpdateEmployee(body) {
  try {
    var data = getSheetValues("employee");
    if (!data) return { status: "error", message: "Sheet 'employee' tidak ditemukan" };
    var headers = data[0];
    var getIdx = function(patterns) {
      return headers.findIndex(function(h) { return patterns.test(String(h).trim()); });
    };
    var emailIdx = headers.findIndex(function(h) { return /email/i.test(String(h).trim()); });
    var userIdx = headers.findIndex(function(h) { return /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()); });
    var idx = {
      name: getIdx(/nama|name|pic/i),
      email: emailIdx !== -1 ? emailIdx : getIdx(/email|user/i),
      username: userIdx !== -1 ? userIdx : -1,
      pos: getIdx(/position|jabatan/i),
      prov: getIdx(/province|provinsi/i),
      area: getIdx(/area/i),
      upline: getIdx(/upline|spv|supervisor|atasan|manager/i),
      password: getIdx(/password|pass/i),
      level: getIdx(/level|grade/i),
      group: getIdx(/group|tim|divisi|division/i)
    };

    if (idx.name === -1) return { status: "error", message: "Kolom nama tidak ditemukan di sheet employee" };

    var targetClean = cleanForMatch(body.originalName);
    var targetRow = -1;
    if (body.originalName) {
      for (var i = 1; i < data.length; i++) {
        if (cleanForMatch(data[i][idx.name]) === targetClean) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (targetRow !== -1) {
      var rowIndex = targetRow - 1;
      if (body.name !== undefined) data[rowIndex][idx.name] = body.name;
      if (idx.username !== -1 && body.user !== undefined) {
        data[rowIndex][idx.username] = body.user;
      } else if (idx.email !== -1 && body.email !== undefined) {
        data[rowIndex][idx.email] = body.email;
      }
      if (idx.pos !== -1 && body.position !== undefined) data[rowIndex][idx.pos] = body.position;
      if (idx.prov !== -1 && body.province !== undefined) data[rowIndex][idx.prov] = body.province;
      if (idx.area !== -1 && body.area !== undefined) data[rowIndex][idx.area] = body.area;
      if (idx.upline !== -1 && body.upline !== undefined) data[rowIndex][idx.upline] = body.upline;
      if (idx.password !== -1 && body.password !== undefined) data[rowIndex][idx.password] = body.password;
      if (idx.level !== -1 && body.level !== undefined) data[rowIndex][idx.level] = body.level;
      if (idx.group !== -1 && body.group !== undefined) data[rowIndex][idx.group] = body.group;
      updateSheetValues("employee", data);
    } else {
      var newRow = [];
      for (var c = 0; c < headers.length; c++) newRow.push("");
      if (idx.name !== -1 && body.name !== undefined) newRow[idx.name] = body.name;
      if (idx.username !== -1 && body.user !== undefined) {
        newRow[idx.username] = body.user;
      } else if (idx.email !== -1 && body.email !== undefined) {
        newRow[idx.email] = body.email;
      }
      if (idx.pos !== -1 && body.position !== undefined) newRow[idx.pos] = body.position;
      if (idx.prov !== -1 && body.province !== undefined) newRow[idx.prov] = body.province;
      if (idx.area !== -1 && body.area !== undefined) newRow[idx.area] = body.area;
      if (idx.upline !== -1 && body.upline !== undefined) newRow[idx.upline] = body.upline;
      if (idx.password !== -1 && body.password !== undefined) newRow[idx.password] = body.password;
      if (idx.level !== -1 && body.level !== undefined) newRow[idx.level] = body.level;
      if (idx.group !== -1 && body.group !== undefined) newRow[idx.group] = body.group;
      
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("employee");
      if (sheet) {
        sheet.appendRow(newRow);
      }
    }
    return { status: "success" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Handle Access Rules Saving
 */
function handleSaveAccessRules(body) {
  try {
    var rules = body.rules || {};
    var headers = ["position", "home", "partner", "stock", "pog", "overview", "temp", "access"];
    var rows = [headers];

    for (var position in rules) {
      var rule = rules[position];
      rows.push([
        position,
        rule.home ? "TRUE" : "FALSE",
        rule.partner ? "TRUE" : "FALSE",
        rule.stock ? "TRUE" : "FALSE",
        rule.pog ? "TRUE" : "FALSE",
        rule.overview ? "TRUE" : "FALSE",
        rule.temp ? "TRUE" : "FALSE",
        rule.access ? "TRUE" : "FALSE"
      ]);
    }

    updateSheetValues("access", rows);
    return { status: "success" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Handle Employee deletion in Apps Script
 */
function handleDeleteEmployee(body) {
  try {
    var data = getSheetValues("employee");
    if (!data) return { status: "error", message: "Sheet 'employee' tidak ditemukan" };
    var headers = data[0];
    var nameIdx = headers.findIndex(function(h) {
      return /nama|name|pic/i.test(String(h).trim());
    });
    if (nameIdx === -1) return { status: "error", message: "Kolom nama employee tidak ditemukan" };

    var targetClean = cleanForMatch(body.name);
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (cleanForMatch(data[i][nameIdx]) === targetClean) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow !== -1) {
      var deletedName = data[targetRow - 1][nameIdx];
      data.splice(targetRow - 1, 1);
      var success = updateSheetValues("employee", data);
      if (!success) return { status: "error", message: "Gagal menyimpan perubahan ke Google Sheets" };
      return { status: "success", message: "Employee " + deletedName + " berhasil dihapus" };
    } else {
      return { status: "error", message: "Employee dengan nama \"" + body.name + "\" tidak ditemukan" };
    }
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Backfill empty category and tgl columns in working sheet
 */
function handleBackfillEmptyCategories(body) {
  try {
    var data = getSheetValues("working");
    if (!data || data.length === 0) {
      return { status: "error", message: "Sheet 'working' tidak ditemukan atau kosong" };
    }

    var headers = data[0];
    
    // Ensure we have Tgl and Category columns
    var hasTgl = false;
    var hasCat = false;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === "tgl" || h === "tanggal") hasTgl = true;
      if (h === "category" || h === "kategori") hasCat = true;
    }

    if (!hasTgl) {
      headers.push("Tgl");
    }
    if (!hasCat) {
      headers.push("Category");
    }

    // Pad each row
    for (var i = 1; i < data.length; i++) {
      while (data[i].length < headers.length) {
        data[i].push("");
      }
    }

    var getIdxList = function(patterns) {
      var indices = [];
      for (var i = 0; i < headers.length; i++) {
        if (patterns.test(String(headers[i]).trim())) {
          indices.push(i);
        }
      }
      return indices;
    };

    var getIdx = function(patterns) {
      return headers.findIndex(function(h) { return patterns.test(String(h).trim()); });
    };

    var timeCols = getIdxList(/^tgl$|^tanggal$|^waktu$|^date$|^timestamp$/i);
    var catCols = getIdxList(/^category$|^kategori$|^klasifikasi$|^cat$/i);
    var idxKiosk = getIdx(/^channel$|^kiosk$/i);

    var now = getJakartaDate();
    var pad = function(n) { return String(n).padStart(2, "0"); };
    var timestamp = pad(now.getDate()) + "/" + pad(now.getMonth() + 1) + "/" + now.getFullYear() + " " + pad(now.getHours()) + ":" + pad(now.getMinutes()) + ":" + pad(now.getSeconds());

    // Load channels
    var channelData = getSheetValues("channel") || [];
    var channelCategoryMap = {};
    if (channelData && channelData.length > 1) {
      var chanHeaders = channelData[0];
      var chanIdxName = chanHeaders.findIndex(function(h) { return /name|nama|channel|kiosk/i.test(String(h).trim()); });
      var chanIdxCat = chanHeaders.findIndex(function(h) { return /kategori|category|klasifikasi/i.test(String(h).trim()); });
      if (chanIdxName !== -1 && chanIdxCat !== -1) {
        for (var c = 1; c < channelData.length; c++) {
          var cRow = channelData[c];
          var cName = String(cRow[chanIdxName] || "").trim().toLowerCase();
          var cCat = String(cRow[chanIdxCat] || "").trim();
          if (cName) {
            channelCategoryMap[cName] = cCat;
          }
        }
      }
    }

    function getGasCategory(kioskName, itemCat) {
      if (itemCat && itemCat !== "-" && itemCat.toLowerCase() !== "uncategorized" && itemCat.toLowerCase() !== "n/a") {
        return itemCat;
      }
      var cleanKName = String(kioskName || "").trim().toLowerCase();
      if (channelCategoryMap[cleanKName]) {
        return channelCategoryMap[cleanKName];
      }
      if (cleanKName.indexOf("distributor") !== -1 || cleanKName.indexOf("distr") !== -1 || cleanKName.indexOf("gudang") !== -1 || cleanKName.indexOf("warehouse") !== -1 || cleanKName.indexOf("subd") !== -1) {
        return "Distributor";
      } else if (/\br1\b|toko r1|mitra r1|kiosk r1/i.test(cleanKName)) {
        return "R1";
      } else if (/\br2\b|toko r2|mitra r2|kiosk r2/i.test(cleanKName)) {
        return "R2";
      }
      return "R1";
    }

    var tglFilled = 0;
    var catFilled = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // 1. Fill empty timestamp/tgl columns
      var hasValidTime = false;
      timeCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
          hasValidTime = true;
        }
      });

      if (!hasValidTime) {
        timeCols.forEach(function(colIdx) {
          row[colIdx] = timestamp;
          tglFilled++;
        });
      } else {
        var existingTimeVal = "";
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "") {
            existingTimeVal = String(row[colIdx]).trim();
          }
        });
        timeCols.forEach(function(colIdx) {
          if (row[colIdx] === undefined || row[colIdx] === null || String(row[colIdx]).trim() === "") {
            row[colIdx] = existingTimeVal;
            tglFilled++;
          }
        });
      }

      // 2. Fill empty category columns
      var kioskNameVal = idxKiosk !== -1 ? String(row[idxKiosk] || "").trim() : "";
      var existingCatVal = "";
      catCols.forEach(function(colIdx) {
        if (row[colIdx] !== undefined && row[colIdx] !== null && String(row[colIdx]).trim() !== "" && String(row[colIdx]).trim().toLowerCase() !== "uncategorized") {
          existingCatVal = String(row[colIdx]).trim();
        }
      });

      if (!existingCatVal) {
        existingCatVal = getGasCategory(kioskNameVal);
        catCols.forEach(function(colIdx) {
          row[colIdx] = existingCatVal;
          catFilled++;
        });
      } else {
        catCols.forEach(function(colIdx) {
          if (row[colIdx] === undefined || row[colIdx] === null || String(row[colIdx]).trim() === "") {
            row[colIdx] = existingCatVal;
            catFilled++;
          }
        });
      }
    }

    var success = updateSheetValues("working", data);
    if (!success) {
      return { status: "error", message: "Gagal menyimpan perubahan ke Google Sheets" };
    }

    return { 
      status: "success", 
      message: "Berhasil mengisi " + catFilled + " kolom category dan " + tglFilled + " kolom tanggal yang kosong di Google Sheets." 
    };

  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}


