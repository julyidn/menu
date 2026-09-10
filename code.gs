/* =====================================================
   KONFIGURASI
   ===================================================== */

const CONFIG = {
  // GANTI DENGAN ID GOOGLE SPREADSHEET ANDA
  SPREADSHEET_ID: '..................',

  // PIN LOGIN APLIKASI
  // Silakan ubah menjadi PIN Anda sendiri.
  PIN: '123456',

  TIMEZONE: 'Asia/Jakarta',

  SHEET_MENU: 'Menu',
  SHEET_GUDANG: 'Gudang Besar',
  SHEET_LOG: 'Log'
};


/* =====================================================
   HEADER SHEET MENU
   ===================================================== */

const HEADER_MENU = [
  'tanggal',

  'makanan_pokok',
  'lauk_hewani',
  'lauk_nabati',
  'sayur',
  'buah',
  'pelengkap',

  'pk_energi',
  'pk_protein',
  'pk_lemak',
  'pk_karbohidrat',
  'pk_serat',

  'pb_energi',
  'pb_protein',
  'pb_lemak',
  'pb_karbohidrat',
  'pb_serat',

  'catatan',

  'foto_id',
  'foto',

  'updated_at'
];


/* =====================================================
   ENTRY POINT GET
   ===================================================== */

function doGet(e) {

  try {

    const params = e && e.parameter ? e.parameter : {};

    const tanggal = normalisasiTanggal(
      params.tanggal || tanggalHariIni()
    );

    return jsonResponse(
      ambilMinggu(tanggal)
    );

  } catch (err) {

    return jsonResponse({
      ok: false,
      pesan: err.message || String(err)
    });

  }

}


/* =====================================================
   ENTRY POINT POST
   ===================================================== */

function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({
        ok: false,
        pesan: 'Data POST kosong.'
      });
    }

    const data = JSON.parse(e.postData.contents);

    const aksi = String(data.aksi || '').trim();

    switch (aksi) {

      case 'masuk':
        return jsonResponse(
          prosesMasuk(data)
        );

      case 'simpan':
        return jsonResponse(
          prosesSimpan(data)
        );

      case 'hapus':
        return jsonResponse(
          prosesHapus(data)
        );

      default:

        return jsonResponse({
          ok: false,
          pesan: 'Aksi tidak dikenal: ' + aksi
        });

    }

  } catch (err) {

    tulisLog(
      'ERROR',
      '',
      err.message || String(err)
    );

    return jsonResponse({
      ok: false,
      pesan: err.message || String(err)
    });

  }

}


/* =====================================================
   LOGIN
   ===================================================== */

function prosesMasuk(data) {

  const pin = String(data.pin || '').trim();

  if (!pin) {

    return {
      ok: false,
      pesan: 'PIN belum diisi.'
    };

  }

  if (pin !== String(CONFIG.PIN)) {

    tulisLog(
      'LOGIN_GAGAL',
      '',
      'PIN salah'
    );

    return {
      ok: false,
      pesan: 'PIN salah.'
    };

  }

  tulisLog(
    'LOGIN',
    '',
    'Login berhasil'
  );

  return {
    ok: true,
    pesan: 'Login berhasil.'
  };

}


/* =====================================================
   SIMPAN MENU
   ===================================================== */

function prosesSimpan(payload) {

  cekPIN(payload.pin);

  if (!payload.data) {

    throw new Error(
      'Data menu tidak ditemukan.'
    );

  }

  const data = payload.data;

  const tanggal = normalisasiTanggal(
    data.tanggal
  );

  if (!tanggal) {

    throw new Error(
      'Tanggal menu tidak valid.'
    );

  }

  if (!data.makanan_pokok) {

    throw new Error(
      'Makanan pokok belum diisi.'
    );

  }


  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const ss = bukaSpreadsheet();

    const sheet = siapkanSheetMenu(ss);

    const semua = sheet.getDataRange().getValues();

    let nomorBaris = -1;

    for (let i = 1; i < semua.length; i++) {

      const tgl = formatTanggal(
        semua[i][0]
      );

      if (tgl === tanggal) {

        nomorBaris = i + 1;
        break;

      }

    }


    /* -----------------------------------------------
       FOTO
       ----------------------------------------------- */

    let fotoId = '';

    let fotoUrl = '';

    const fotoLama = String(
      data.fotoLama || ''
    ).trim();

    const fotoBaru = String(
      data.fotoBaru || ''
    ).trim();

    const hapusFoto =
      data.hapusFoto === true;


    if (fotoBaru) {

      const hasilFoto = simpanFoto(
        fotoBaru,
        tanggal
      );

      fotoId = hasilFoto.id;
      fotoUrl = hasilFoto.url;

      // Hapus foto lama jika ada
      if (fotoLama) {

        hapusFileDriveAman(
          fotoLama
        );

      }

    } else if (hapusFoto) {

      if (fotoLama) {

        hapusFileDriveAman(
          fotoLama
        );

      }

      fotoId = '';
      fotoUrl = '';

    } else {

      // Pertahankan foto lama
      if (nomorBaris !== -1) {

        fotoId =
          semua[nomorBaris - 1][18] || '';

        fotoUrl =
          semua[nomorBaris - 1][19] || '';

      }

    }


    /* -----------------------------------------------
       DATA MENU
       ----------------------------------------------- */

    const row = buatRowMenu(
      data,
      tanggal,
      fotoId,
      fotoUrl
    );


    if (nomorBaris === -1) {

      sheet.appendRow(row);

    } else {

      sheet
        .getRange(
          nomorBaris,
          1,
          1,
          HEADER_MENU.length
        )
        .setValues([row]);

    }


    tulisLog(
      'SIMPAN',
      tanggal,
      'Menu disimpan'
    );


    return {

      ok: true,

      pesan: 'Menu berhasil disimpan.',

      tanggal: tanggal

    };


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
   HAPUS MENU
   ===================================================== */

function prosesHapus(payload) {

  cekPIN(payload.pin);

  const tanggal =
    normalisasiTanggal(
      payload.tanggal
    );

  if (!tanggal) {

    throw new Error(
      'Tanggal tidak valid.'
    );

  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const ss = bukaSpreadsheet();

    const sheet =
      siapkanSheetMenu(ss);

    const semua =
      sheet.getDataRange().getValues();


    for (let i = semua.length - 1; i >= 1; i--) {

      const tgl =
        formatTanggal(
          semua[i][0]
        );

      if (tgl === tanggal) {

        const fotoId =
          semua[i][18] || '';

        if (fotoId) {

          hapusFileDriveAman(
            fotoId
          );

        }

        sheet.deleteRow(
          i + 1
        );

        tulisLog(
          'HAPUS',
          tanggal,
          'Menu dihapus'
        );

        return {

          ok: true,

          pesan:
            'Menu berhasil dihapus.',

          tanggal: tanggal

        };

      }

    }


    return {

      ok: true,

      pesan:
        'Menu tidak ditemukan.',

      tanggal: tanggal

    };


  } finally {

    lock.releaseLock();

  }

}


/* =====================================================
   AMBIL DATA MINGGU
   ===================================================== */

function ambilMinggu(tanggal) {

  const senin =
    awalMinggu(tanggal);

  const hasil = [];

  const ss =
    bukaSpreadsheet();

  const sheet =
    siapkanSheetMenu(ss);

  const semua =
    sheet.getDataRange().getValues();


  const map = {};


  for (let i = 1; i < semua.length; i++) {

    const row =
      semua[i];

    const tgl =
      formatTanggal(row[0]);

    if (!tgl) continue;

    map[tgl] =
      rowMenuKeObject(row);

  }


  for (let i = 0; i < 7; i++) {

    const tgl =
      tambahHari(
        senin,
        i
      );

    if (map[tgl]) {

      hasil.push(
        map[tgl]
      );

    } else {

      hasil.push({

        tanggal: tgl,

        makanan_pokok: '',
        lauk_hewani: '',
        lauk_nabati: '',
        sayur: '',
        buah: '',
        pelengkap: '',

        gizi: {
          pk: {},
          pb: {}
        },

        catatan: '',

        foto_id: '',
        foto: ''

      });

    }

  }


  return {

    ok: true,

    minggu: hasil

  };

}


/* =====================================================
   CONVERT ROW -> OBJECT
   ===================================================== */

function rowMenuKeObject(row) {

  return {

    tanggal:
      formatTanggal(row[0]),

    makanan_pokok:
      safeString(row[1]),

    lauk_hewani:
      safeString(row[2]),

    lauk_nabati:
      safeString(row[3]),

    sayur:
      safeString(row[4]),

    buah:
      safeString(row[5]),

    pelengkap:
      safeString(row[6]),


    gizi: {

      pk: {

        energi:
          safeValue(row[7]),

        protein:
          safeValue(row[8]),

        lemak:
          safeValue(row[9]),

        karbohidrat:
          safeValue(row[10]),

        serat:
          safeValue(row[11])

      },

      pb: {

        energi:
          safeValue(row[12]),

        protein:
          safeValue(row[13]),

        lemak:
          safeValue(row[14]),

        karbohidrat:
          safeValue(row[15]),

        serat:
          safeValue(row[16])

      }

    },


    catatan:
      safeString(row[17]),

    foto_id:
      safeString(row[18]),

    foto:
      safeString(row[19])

  };

}


/* =====================================================
   OBJECT -> ROW
   ===================================================== */

function buatRowMenu(
  data,
  tanggal,
  fotoId,
  fotoUrl
) {

  return [

    tanggal,

    safeString(
      data.makanan_pokok
    ),

    safeString(
      data.lauk_hewani
    ),

    safeString(
      data.lauk_nabati
    ),

    safeString(
      data.sayur
    ),

    safeString(
      data.buah
    ),

    safeString(
      data.pelengkap
    ),


    safeNumber(
      data.pk_energi
    ),

    safeNumber(
      data.pk_protein
    ),

    safeNumber(
      data.pk_lemak
    ),

    safeNumber(
      data.pk_karbohidrat
    ),

    safeNumber(
      data.pk_serat
    ),


    safeNumber(
      data.pb_energi
    ),

    safeNumber(
      data.pb_protein
    ),

    safeNumber(
      data.pb_lemak
    ),

    safeNumber(
      data.pb_karbohidrat
    ),

    safeNumber(
      data.pb_serat
    ),


    safeString(
      data.catatan
    ),


    fotoId,

    fotoUrl,

    new Date()

  ];

}


/* =====================================================
   SHEET MENU
   ===================================================== */

function siapkanSheetMenu(ss) {

  let sheet =
    ss.getSheetByName(
      CONFIG.SHEET_MENU
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.SHEET_MENU
      );

  }


  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        HEADER_MENU.length
      )
      .setValues([
        HEADER_MENU
      ]);

    sheet
      .getRange(
        1,
        1,
        1,
        HEADER_MENU.length
      )
      .setFontWeight(
        'bold'
      );

  }


  return sheet;

}


/* =====================================================
   GUDANG BESAR
   ===================================================== */

/*
 Format:

 Kolom A = Nama Bahan
 Kolom B = Satuan
 Kolom C = Nama Alias

 Contoh:

 Beras | kg | beras putih,nasi,beras premium
 Ayam  | kg | ayam potong,daging ayam
 Tempe | kg | tempe kedelai
*/


function siapkanGudang() {

  const ss =
    bukaSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.SHEET_GUDANG
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.SHEET_GUDANG
      );

    sheet
      .getRange(
        1,
        1,
        1,
        3
      )
      .setValues([[
        'Nama Bahan',
        'Satuan',
        'Nama Alias'
      ]]);

    sheet
      .getRange(
        1,
        1,
        1,
        3
      )
      .setFontWeight(
        'bold'
      );

  }


  return sheet;

}


/* =====================================================
   PENCARIAN MASTER GUDANG
   ===================================================== */

function cariBahanMaster(
  namaCSV
) {

  const sheet =
    siapkanGudang();

  const values =
    sheet.getDataRange()
      .getValues();


  const target =
    normalisasiNama(
      namaCSV
    );


  if (!target) {
    return null;
  }


  /*
   * 1. Cari nama persis
   */

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const nama =
      normalisasiNama(
        values[i][0]
      );

    if (
      nama &&
      nama === target
    ) {

      return {

        nama:
          String(values[i][0]),

        satuan:
          String(values[i][1] || ''),

        skor: 1,

        baris: i + 1

      };

    }

  }


  /*
   * 2. Cari alias
   */

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const aliases =
      String(
        values[i][2] || ''
      )
      .split(',')
      .map(function(x) {
        return normalisasiNama(x);
      })
      .filter(Boolean);


    if (
      aliases.indexOf(target) !== -1
    ) {

      return {

        nama:
          String(values[i][0]),

        satuan:
          String(values[i][1] || ''),

        skor: 0.95,

        baris: i + 1

      };

    }

  }


  /*
   * 3. Fuzzy sederhana
   */

  let terbaik = null;

  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const nama =
      normalisasiNama(
        values[i][0]
      );

    if (!nama) continue;


    const skor =
      skorKemiripan(
        target,
        nama
      );


    if (
      skor >= 0.72 &&
      (!terbaik ||
       skor > terbaik.skor)
    ) {

      terbaik = {

        nama:
          String(values[i][0]),

        satuan:
          String(values[i][1] || ''),

        skor: skor,

        baris: i + 1

      };

    }

  }


  return terbaik;

}


/* =====================================================
   FUZZY MATCHING
   ===================================================== */

function skorKemiripan(
  a,
  b
) {

  if (!a || !b) {
    return 0;
  }


  if (a === b) {
    return 1;
  }


  if (
    a.indexOf(b) !== -1 ||
    b.indexOf(a) !== -1
  ) {

    const panjang =
      Math.max(
        a.length,
        b.length
      );

    const pendek =
      Math.min(
        a.length,
        b.length
      );

    return
      pendek / panjang;

  }


  const kataA =
    a.split(' ');

  const kataB =
    b.split(' ');


  let cocok = 0;

  kataA.forEach(
    function(kata) {

      if (
        kata.length >= 3 &&
        kataB.indexOf(kata) !== -1
      ) {

        cocok++;

      }

    }
  );


  return cocok /
    Math.max(
      kataA.length,
      kataB.length
    );

}


/* =====================================================
   NORMALISASI NAMA
   ===================================================== */

function normalisasiNama(
  teks
) {

  return String(
    teks || ''
  )
  .toLowerCase()
  .normalize('NFD')
  .replace(
    /[\u0300-\u036f]/g,
    ''
  )
  .replace(
    /[^a-z0-9\s]/g,
    ' '
  )
  .replace(
    /\s+/g,
    ' '
  )
  .trim();

}


/* =====================================================
   NORMALISASI TANGGAL
   ===================================================== */

function normalisasiTanggal(
  nilai
) {

  if (!nilai) {
    return '';
  }


  if (
    Object.prototype.toString
      .call(nilai) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      nilai,
      CONFIG.TIMEZONE,
      'yyyy-MM-dd'
    );

  }


  const teks =
    String(nilai)
      .trim();


  /*
   * Sudah YYYY-MM-DD
   */

  if (
    /^\d{4}-\d{2}-\d{2}$/
      .test(teks)
  ) {

    return teks;

  }


  /*
   * DD/MM/YYYY
   */

  let m =
    teks.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
    );

  if (m) {

    return [
      m[3],
      pad2(m[2]),
      pad2(m[1])
    ].join('-');

  }


  /*
   * DD-MM-YYYY
   */

  m =
    teks.match(
      /^(\d{1,2})-(\d{1,2})-(\d{4})$/
    );

  if (m) {

    return [
      m[3],
      pad2(m[2]),
      pad2(m[1])
    ].join('-');

  }


  return '';

}


/* =====================================================
   TANGGAL
   ===================================================== */

function tanggalHariIni() {

  return Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    'yyyy-MM-dd'
  );

}


function awalMinggu(
  tanggal
) {

  const parts =
    tanggal.split('-');


  const d =
    new Date(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2])
    );


  const hari =
    d.getDay();


  const selisih =
    (hari + 6) % 7;


  d.setDate(
    d.getDate() - selisih
  );


  return formatDateLocal(d);

}


function tambahHari(
  tanggal,
  jumlah
) {

  const parts =
    tanggal.split('-');


  const d =
    new Date(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2])
    );


  d.setDate(
    d.getDate() + jumlah
  );


  return formatDateLocal(d);

}


function formatDateLocal(
  d
) {

  return [
    d.getFullYear(),
    pad2(
      d.getMonth() + 1
    ),
    pad2(
      d.getDate()
    )
  ].join('-');

}


/* =====================================================
   FOTO
   ===================================================== */

function simpanFoto(
  dataUrl,
  tanggal
) {

  /*
   * Folder otomatis dibuat.
   */

  const folder =
    cariAtauBuatFolder(
      'Foto Menu SPPG'
    );


  const match =
    dataUrl.match(
      /^data:(image\/[^;]+);base64,(.+)$/
    );


  if (!match) {

    throw new Error(
      'Format foto tidak valid.'
    );

  }


  const mime =
    match[1];

  const bytes =
    Utilities.base64Decode(
      match[2]
    );


  const ekstensi =
    mime.indexOf('png') !== -1
      ? 'png'
      : 'jpg';


  const namaFile =
    'Menu-' +
    tanggal +
    '-' +
    new Date().getTime() +
    '.' +
    ekstensi;


  const blob =
    Utilities.newBlob(
      bytes,
      mime,
      namaFile
    );


  const file =
    folder.createFile(
      blob
    );


  /*
   * Supaya gambar bisa ditampilkan
   * oleh frontend.
   */

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );


  const id =
    file.getId();


  const url =
    'https://drive.google.com/uc?export=view&id=' +
    encodeURIComponent(id);


  return {

    id: id,

    url: url

  };

}


/* =====================================================
   FOLDER DRIVE
   ===================================================== */

function cariAtauBuatFolder(
  nama
) {

  const iterator =
    DriveApp.getFoldersByName(
      nama
    );


  if (
    iterator.hasNext()
  ) {

    return iterator.next();

  }


  return DriveApp.createFolder(
    nama
  );

}


/* =====================================================
   HAPUS FOTO
   ===================================================== */

function hapusFileDriveAman(
  fileId
) {

  try {

    if (!fileId) {
      return;
    }

    const file =
      DriveApp.getFileById(
        fileId
      );

    file.setTrashed(
      true
    );

  } catch (err) {

    /*
     * Jangan membuat penyimpanan
     * menu gagal hanya karena file
     * foto lama sudah tidak tersedia.
     */

    tulisLog(
      'FOTO_ERROR',
      '',
      err.message || String(err)
    );

  }

}


/* =====================================================
   SPREADSHEET
   ===================================================== */

function bukaSpreadsheet() {

  if (
    !CONFIG.SPREADSHEET_ID ||
    CONFIG.SPREADSHEET_ID ===
      'GANTI_DENGAN_ID_SPREADSHEET'
  ) {

    throw new Error(
      'SPREADSHEET_ID belum diisi.'
    );

  }


  return SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

}


/* =====================================================
   LOG
   ===================================================== */

function tulisLog(
  aksi,
  tanggal,
  keterangan
) {

  try {

    const ss =
      bukaSpreadsheet();

    let sheet =
      ss.getSheetByName(
        CONFIG.SHEET_LOG
      );


    if (!sheet) {

      sheet =
        ss.insertSheet(
          CONFIG.SHEET_LOG
        );

      sheet
        .getRange(
          1,
          1,
          1,
          4
        )
        .setValues([[
          'Waktu',
          'Aksi',
          'Tanggal',
          'Keterangan'
        ]]);

    }


    sheet.appendRow([

      new Date(),

      aksi,

      tanggal,

      keterangan

    ]);

  } catch (err) {

    console.log(
      'Log error:',
      err
    );

  }

}


/* =====================================================
   CEK PIN
   ===================================================== */

function cekPIN(
  pin
) {

  const nilai =
    String(pin || '')
      .trim();


  if (
    !nilai ||
    nilai !== String(CONFIG.PIN)
  ) {

    throw new Error(
      'PIN salah atau sesi tidak valid.'
    );

  }

}


/* =====================================================
   RESPONSE JSON
   ===================================================== */

function jsonResponse(
  object
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(object)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


/* =====================================================
   HELPER
   ===================================================== */

function safeString(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  return String(value);

}


function safeNumber(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return '';

  }


  const n =
    Number(
      String(value)
        .replace(',', '.')
    );


  return isNaN(n)
    ? ''
    : n;

}


function safeValue(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  return value;

}


function formatTanggal(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return '';

  }


  if (
    Object.prototype.toString
      .call(value) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      CONFIG.TIMEZONE,
      'yyyy-MM-dd'
    );

  }


  const teks =
    String(value)
      .trim();


  if (
    /^\d{4}-\d{2}-\d{2}$/
      .test(teks)
  ) {

    return teks;

  }


  return normalisasiTanggal(
    teks
  );

}


function pad2(
  value
) {

  return String(
    value
  ).padStart(
    2,
    '0'
  );

}


/* =====================================================
   TEST LOGIN
   ===================================================== */

function testLogin() {

  const hasil =
    prosesMasuk({
      pin: CONFIG.PIN
    });


  Logger.log(
    JSON.stringify(
      hasil
    )
  );

}


/* =====================================================
   TEST MINGGU
   ===================================================== */

function testMinggu() {

  const hasil =
    ambilMinggu(
      tanggalHariIni()
    );


  Logger.log(
    JSON.stringify(
      hasil,
      null,
      2
    )
  );

}


/* =====================================================
   SETUP AWAL
   ===================================================== */

function setup() {

  const ss =
    bukaSpreadsheet();


  /*
   * Buat Menu
   */

  siapkanSheetMenu(
    ss
  );


  /*
   * Buat Gudang Besar
   */

  siapkanGudang();


  /*
   * Buat Log
   */

  tulisLog(
    'SETUP',
    '',
    'Backend berhasil disiapkan.'
  );


  Logger.log(
    'Setup selesai.'
  );

}