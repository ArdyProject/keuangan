const SPREADSHEET_ID = '1FpTEV2aJyyVtWt30rny_rWvrYni8pE0dIDuHzr8D02w';
const SHEET_NAME = 'Sheet1';

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan.");
  return sheet;
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fungsi yang sudah diperbaiki agar lebih fleksibel mendeteksi jenis transaksi
function normalizeJenis_(value) {
  const jenis = String(value == null ? '' : value).trim().toLowerCase();

  const keywordsPemasukan = ['+', 'pemasukan', 'masuk', 'income', 'in', 'debit'];
  const keywordsPengeluaran = ['-', 'pengeluaran', 'keluar', 'expense', 'out', 'kredit'];

  if (keywordsPemasukan.some(keyword => jenis.includes(keyword))) {
    return 'Pemasukan';
  }

  if (keywordsPengeluaran.some(keyword => jenis.includes(keyword))) {
    return 'Pengeluaran';
  }

  return null;
}

function parseNominal_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;

  let text = String(value == null ? '' : value)
    .trim()
    .replace(/Rp/gi, '')
    .replace(/\s/g, '');

  if (!text) return 0;

  if (text.includes(',')) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else {
    text = text.replace(/\./g, '');
  }

  const number = Number(text.replace(/[^0-9.-]/g, ''));
  return isFinite(number) ? number : 0;
}

function formatTanggal_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(value == null ? '' : value).trim();
}

function doGet() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) return jsonOutput_([]);

    const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const data = rows
      .filter(row => String(row[0] == null ? '' : row[0]).trim() !== '')
      .map(row => ({
        id: String(row[0]),
        tanggal: formatTanggal_(row[1]),
        jenis: normalizeJenis_(row[2]) || '',
        nominal: parseNominal_(row[3]),
        keterangan: String(row[4] == null ? '' : row[4])
      }));

    return jsonOutput_(data);
  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request POST tidak berisi data.');
    }

    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const action = String(data.action || '').trim().toLowerCase();

    if (action === 'add') {
      const tanggal = String(data.tanggal || '').trim();
      const jenis = normalizeJenis_(data.jenis);
      const nominal = parseNominal_(data.nominal);
      const keterangan = String(data.keterangan || '').trim();

      if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) {
        throw new Error('Tanggal harus berformat yyyy-MM-dd.');
      }
      if (!jenis) {
        throw new Error('Jenis harus Pemasukan atau Pengeluaran.');
      }
      if (nominal <= 0) {
        throw new Error('Nominal harus lebih besar dari 0.');
      }
      if (!keterangan) {
        throw new Error('Keterangan wajib diisi.');
      }

      const id = Utilities.getUuid();
      const row = sheet.getLastRow() + 1;

      // Kolom C sengaja menyimpan kata, bukan + atau -, agar tidak pernah dianggap formula.
      sheet.getRange(row, 1, 1, 5).setValues([[
        id,
        tanggal,
        jenis,
        nominal,
        keterangan
      ]]);

      return jsonOutput_({ status: 'success', id: id });
    }

    if (action === 'delete') {
      const wantedId = String(data.id || '').trim();
      if (!wantedId) throw new Error('ID transaksi kosong.');

      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === wantedId) {
            sheet.deleteRow(i + 2);
            return jsonOutput_({ status: 'success' });
          }
        }
      }

      return jsonOutput_({
        status: 'error',
        message: 'ID tidak ditemukan.'
      });
    }

    return jsonOutput_({
      status: 'error',
      message: 'Action tidak dikenal.'
    });
  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  }
}