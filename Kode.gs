function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DOMPET RUMAH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SETUP DATABASE
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Transaksi': ['ID', 'Tanggal', 'Tipe', 'Kategori', 'Nominal', 'Dompet', 'Keterangan'],
    'Dompet': ['ID_Dompet', 'Nama_Dompet', 'Tipe', 'Saldo'],
    'Budget': ['Kategori', 'Batas_Maksimal', 'Bulan'],
    'Settings': ['Kunci', 'Nilai']
  };

  for (const [sheetName, headers] of Object.entries(sheets)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      if(sheetName === 'Settings') {
        sheet.appendRow(['PIN', '1234']);
        sheet.appendRow(['Dompet', 'Kas Tunai,Rek. BCA,Rek. Mandiri,GoPay']);
        sheet.appendRow(['KatPengeluaran', 'Makan & Minum,Belanja Bulanan,Listrik & Air,Pendidikan,Hiburan']);
        // Menambahkan Saldo Awal sebagai default pemasukan
        sheet.appendRow(['KatPemasukan', 'Saldo Awal,Gaji Utama,Bonus,Hasil Usaha,Lainnya']);
      }
    }
  }
  return "Database berhasil disiapkan!";
}

// ==========================================
// KALKULASI KEKAYAAN BERSIH & INISIALISASI
// ==========================================
function getNetWorth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transaksi');
  if(!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if(data.length <= 1) return null; // Belum ada transaksi sama sekali
  
  let totalIn = 0;
  let totalOut = 0;
  
  for(let i = 1; i < data.length; i++) {
    let nominal = Number(data[i][4]) || 0;
    if(data[i][2] === 'Pemasukan') totalIn += nominal;
    if(data[i][2] === 'Pengeluaran') totalOut += nominal;
  }
  
  return totalIn - totalOut;
}

// Mengambil pengaturan dan saldo dalam satu kali panggil agar ringan
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  if(!sheet) { setupDatabase(); sheet = ss.getSheetByName('Settings'); }
  
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i = 1; i < data.length; i++) { settings[data[i][0]] = data[i][1]; }
  
  return {
    settings: settings,
    netWorth: getNetWorth()
  };
}

function saveSettings(newSettings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      let key = data[i][0];
      if(newSettings[key] !== undefined) sheet.getRange(i + 1, 2).setValue(newSettings[key]);
    }
    return { success: true, message: "Pengaturan berhasil diperbarui!" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// FITUR TRANSAKSI & TRANSFER
// ==========================================
function addTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Transaksi');
    const id = "TRX-" + new Date().getTime();
    sheet.appendRow([id, data.tanggal, data.tipe, data.kategori, data.nominal, data.dompet, data.keterangan]);
    return { success: true, message: "Transaksi dicatat", newNetWorth: getNetWorth() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function addTransfer(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Transaksi');
    const timeId = new Date().getTime();
    sheet.appendRow(["TRF-OUT-" + timeId, data.tanggal, "Pengeluaran", "Transfer Keluar", data.nominal, data.dariDompet, "Transfer ke " + data.keDompet + " (" + data.keterangan + ")"]);
    sheet.appendRow(["TRF-IN-" + timeId, data.tanggal, "Pemasukan", "Transfer Masuk", data.nominal, data.keDompet, "Transfer dari " + data.dariDompet + " (" + data.keterangan + ")"]);
    return { success: true, message: "Transfer berhasil", newNetWorth: getNetWorth() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// FITUR LAPORAN & BUDGET
// ==========================================
function getReportData(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName('Transaksi');
    if(!sheet) return { success: false, message: "Sheet Transaksi tidak ditemukan" };
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { success: true, data: [], summary: {totalIn:0, totalOut:0} };
    data.shift(); let filteredData = [], totalIn = 0, totalOut = 0;
    const start = startDate ? new Date(startDate) : new Date(0); start.setHours(0,0,0,0);
    const end = endDate ? new Date(endDate) : new Date(); end.setHours(23, 59, 59, 999);
    
    data.forEach(row => {
      let trxDate; let dateStr = row[1].toString();
      if(dateStr.includes('/')) { let parts = dateStr.split('/'); trxDate = new Date(parts[2], parts[1] - 1, parts[0]); } else { trxDate = new Date(dateStr); }
      if (trxDate >= start && trxDate <= end) {
        filteredData.push({ id: row[0], tanggal: row[1], tipe: row[2], kategori: row[3], nominal: row[4], dompet: row[5], keterangan: row[6] });
        if(row[2] === 'Pemasukan') totalIn += Number(row[4]);
        if(row[2] === 'Pengeluaran') totalOut += Number(row[4]);
      }
    });
    filteredData.reverse();
    return { success: true, data: filteredData, summary: { totalIn: totalIn, totalOut: totalOut } };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function saveBudget(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName('Budget');
    if(!sheet) { sheet = ss.insertSheet('Budget'); sheet.appendRow(['Kategori', 'Batas_Maksimal', 'Bulan']); }
    const dataRange = sheet.getDataRange().getValues(); let found = false;
    for(let i = 1; i < dataRange.length; i++) { if(dataRange[i][0] === data.kategori) { sheet.getRange(i + 1, 2).setValue(data.batas); found = true; break; } }
    if(!found) { sheet.appendRow([data.kategori, data.batas, new Date().getMonth() + 1]); }
    return { success: true, message: "Anggaran disimpan!" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getBudgetStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBudget = ss.getSheetByName('Budget'); const sheetTrx = ss.getSheetByName('Transaksi');
  if(!sheetBudget || !sheetTrx) return [];
  const budgets = sheetBudget.getDataRange().getValues(); const trxs = sheetTrx.getDataRange().getValues();
  if(budgets.length <= 1) return [];
  const currentMonth = new Date().getMonth(); const currentYear = new Date().getFullYear();
  let expenses = {};
  for(let i = 1; i < trxs.length; i++) {
    const tipe = trxs[i][2];
    if(tipe === 'Pengeluaran') {
      const kat = trxs[i][3]; const dateStr = trxs[i][1].toString();
      let trxDate; if(dateStr.includes('/')) { let parts = dateStr.split('/'); trxDate = new Date(parts[2], parts[1] - 1, parts[0]); } else { trxDate = new Date(dateStr); }
      if(trxDate.getMonth() === currentMonth && trxDate.getFullYear() === currentYear) { expenses[kat] = (expenses[kat] || 0) + Number(trxs[i][4]); }
    }
  }
  let status = [];
  for(let i = 1; i < budgets.length; i++) {
    const kat = budgets[i][0]; const max = budgets[i][1]; const spent = expenses[kat] || 0;
    const percentage = max > 0 ? (spent / max) * 100 : 0;
    status.push({ kategori: kat, maksimal: max, terpakai: spent, persentase: percentage });
  }
  return status;
}