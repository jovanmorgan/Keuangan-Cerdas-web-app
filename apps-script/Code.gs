/**
 * KEUANGAN CERDAS - Backend (Google Apps Script, container-bound ke Spreadsheet)
 */
const SH_TRX = "Transaksi",
  SH_KAT = "Kategori",
  SH_SET = "Pengaturan",
  SH_ANG = "Anggaran",
  SH_TGT = "Target";
const DEF_SET_ = {
  MIN_SALDO_WARNING: 50000,
  TARGET_TABUNGAN: 20,
  BUDGET_TOTAL: 0,
  ALERT_BUDGET_PCT: 80,
  LONJAKAN_PCT: 30,
  BATAS_TRX_BESAR: 500000,
  HARI_TANPA_CATAT: 3,
  DANA_DARURAT_BULAN: 3,
  ALERT_POPUP: 1,
};

function doGet(e) {
  // Dipanggil sebagai API (dari browser bar / testing) lewat query string, mis:
  // .../exec?fn=getAllData&args=[]&token=xxx
  const p = (e && e.parameter) || {};
  if (p.fn) {
    let args = [];
    try {
      args = p.args ? JSON.parse(p.args) : [];
    } catch (_) {
      args = [];
    }
    return handleApi_(p.fn, args, p.token);
  }
  setupDatabase();
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Keuangan Cerdas")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Dipanggil dari frontend yang dihosting terpisah (mis. GitHub + Vercel) lewat
 * fetch() sebagai POST dengan Content-Type: text/plain berisi JSON:
 *   { "fn": "namaFungsi", "args": [...], "token": "TOKEN_RAHASIA" }
 * Content-Type text/plain sengaja dipakai supaya browser tidak melakukan CORS
 * preflight (OPTIONS), karena Apps Script Web App tidak mendukungnya.
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    return handleApi_(body.fn, body.args, body.token);
  } catch (err) {
    return jsonOutput_(err_("Permintaan tidak valid: " + err.message));
  }
}

/**
 * Daftar fungsi yang boleh dipanggil dari luar lewat doGet/doPost. Ini adalah
 * whitelist eksplisit; fungsi apa pun di luar daftar ini TIDAK bisa dipanggil
 * lewat web API meskipun namanya ditebak dengan benar.
 */
const API_FUNCTIONS_ = {
  getAllData: getAllData,
  simpanDataTransaksi: simpanDataTransaksi,
  simpanBanyakTransaksi: simpanBanyakTransaksi,
  hapusTransaksi: hapusTransaksi,
  hapusTransaksiBanyak: hapusTransaksiBanyak,
  editTransaksiBanyak: editTransaksiBanyak,
  simpanKategori: simpanKategori,
  hapusKategori: hapusKategori,
  simpanPengaturan: simpanPengaturan,
  simpanAnggaran: simpanAnggaran,
  hapusAnggaran: hapusAnggaran,
  simpanTarget: simpanTarget,
  hapusTarget: hapusTarget,
  setorTarget: setorTarget,
  analisisSuaraPintar: analisisSuaraPintar,
};

/**
 * Token rahasia (opsional tapi SANGAT disarankan) supaya orang lain yang
 * kebetulan tahu URL Web App tidak bisa ikut mengubah data Sheet.
 * Atur di: Project Settings (⚙️) > Script Properties > tambah key "APP_TOKEN".
 * Kalau APP_TOKEN belum diatur, semua permintaan akan diterima tanpa token
 * (memudahkan waktu pengembangan, tapi tidak disarankan untuk pemakaian nyata).
 */
function checkToken_(token) {
  const required = PropertiesService.getScriptProperties().getProperty(
    "APP_TOKEN",
  );
  if (!required) return true;
  return String(token || "") === String(required);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function handleApi_(fn, args, token) {
  if (!checkToken_(token)) return jsonOutput_(err_("Token tidak valid"));
  if (!fn || !API_FUNCTIONS_.hasOwnProperty(fn))
    return jsonOutput_(err_("Fungsi tidak dikenali: " + fn));
  try {
    setupDatabase();
    const result = API_FUNCTIONS_[fn].apply(null, Array.isArray(args) ? args : []);
    return jsonOutput_(result === undefined ? ok_("") : result);
  } catch (err) {
    return jsonOutput_(err_("Gagal memproses permintaan: " + err.message));
  }
}

/* ---------- Helper ---------- */
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    setupDatabase();
    s = ss.getSheetByName(name);
  }
  return s;
}
const ok_ = (m) => ({ status: "success", message: m });
const err_ = (m) => ({ status: "error", message: m });

/** Jalankan operasi tulis dengan kunci agar data tidak bentrok. */
function tx_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const r = fn();
    SpreadsheetApp.flush();
    return r;
  } catch (e) {
    return err_("Gagal: " + e.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

// Sheets mengubah teks tanggal jadi Date; Date tidak bisa dikirim ke frontend, jadi diformat ulang.
function fmtTanggal_(v) {
  return v instanceof Date
    ? Utilities.formatDate(
        v,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss",
      )
    : String(v || "");
}

/** Cari nomor baris (tanpa membedakan huruf besar/kecil) berdasarkan nilai di kolom tertentu. */
function findBy_(s, col, val) {
  const last = s.getLastRow();
  if (last < 2) return 0;
  const v = s.getRange(2, col, last - 1, 1).getValues(),
    t = String(val).toLowerCase();
  for (let i = 0; i < v.length; i++)
    if (String(v[i][0]).toLowerCase() === t) return i + 2;
  return 0;
}
function rows_(name, cols) {
  const s = getSheet_(name),
    l = s.getLastRow();
  return l > 1
    ? s
        .getRange(2, 1, l - 1, cols)
        .getValues()
        .filter((r) => r[0] !== "")
    : [];
}
function getSettings_() {
  const o = Object.assign({}, DEF_SET_);
  rows_(SH_SET, 2).forEach((r) => {
    if (r[0] in DEF_SET_ && r[1] !== "" && !isNaN(r[1])) o[r[0]] = Number(r[1]);
  });
  return o;
}

/**
 * Sinkronkan kategori yang "nyelip" di sheet Anggaran / Transaksi (misalnya diketik
 * manual langsung di Sheet, bukan lewat form di aplikasi) supaya otomatis
 * terdaftar juga di sheet Kategori.
 */
function syncKategoriDariSheetLain_() {
  const kat = getSheet_(SH_KAT);
  const existing = new Set(
    rows_(SH_KAT, 3)
      .map((r) => String(r[1] || "").toLowerCase())
      .filter(Boolean),
  );
  let nextId = Math.max(0, kat.getLastRow() - 1);
  const toAdd = [];

  const daftarkan_ = (nama, tipe) => {
    nama = String(nama || "").trim();
    const key = nama.toLowerCase();
    if (!nama || existing.has(key)) return;
    existing.add(key);
    nextId++;
    toAdd.push([nextId, nama, tipe]);
  };

  // Kategori yang dipakai di Anggaran tapi belum ada di Kategori -> anggap Pengeluaran
  rows_(SH_ANG, 2).forEach((r) => daftarkan_(r[0], "Pengeluaran"));

  // Kategori yang dipakai di Transaksi tapi belum ada di Kategori -> ikuti jenis transaksinya
  rows_(SH_TRX, 7).forEach((r) =>
    daftarkan_(r[4], r[3] === "Pemasukan" ? "Pemasukan" : "Pengeluaran"),
  );

  if (toAdd.length)
    kat.getRange(kat.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
}

/* ---------- Baca semua data (dipanggil frontend tiap beberapa detik) ---------- */
function getAllData() {
  tx_(() => {
    syncKategoriDariSheetLain_();
    return ok_("");
  });
  const trx = rows_(SH_TRX, 7).map((r) => ({
    id: String(r[0]),
    tanggal: fmtTanggal_(r[1]),
    nama: String(r[2] || "-"),
    jenis: r[3] === "Pemasukan" ? "Pemasukan" : "Pengeluaran",
    kategori: String(r[4] || "Umum"),
    nominal: Number(r[5]) || 0,
    keterangan: String(r[6] || ""),
  }));
  trx.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  return {
    trx: trx,
    kategori: rows_(SH_KAT, 3)
      .map((r) => ({ nama: String(r[1]), tipe: String(r[2]) }))
      .filter((k) => k.nama),
    anggaran: rows_(SH_ANG, 2).map((r) => ({
      kategori: String(r[0]),
      batas: Number(r[1]) || 0,
    })),
    target: rows_(SH_TGT, 5).map((r) => ({
      id: String(r[0]),
      nama: String(r[1]),
      target: Number(r[2]) || 0,
      terkumpul: Number(r[3]) || 0,
      tenggat: fmtTanggal_(r[4]).slice(0, 10),
    })),
    settings: getSettings_(),
  };
}

/* ---------- Transaksi ---------- */
function simpanDataTransaksi(f) {
  return tx_(() => {
    const nama = String(f.nama || "").trim(),
      nominal = Number(f.nominal);
    if (!nama || !(nominal > 0))
      return err_("Nama dan nominal (lebih dari 0) wajib diisi");
    const jenis = f.jenis === "Pemasukan" ? "Pemasukan" : "Pengeluaran";
    const kategori = String(f.kategori || "").trim() || "Umum";
    const tgl = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(f.tanggal || "")
      ? f.tanggal
      : Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm:ss",
        );
    const s = getSheet_(SH_TRX);
    let row = f.id ? findBy_(s, 1, f.id) : 0;
    if (f.id && !row)
      return err_("Data tidak ditemukan, mungkin sudah dihapus");
    const isEdit = row > 0;
    if (!isEdit) row = s.getLastRow() + 1;
    s.getRange(row, 2).setNumberFormat("@");
    s.getRange(row, 1, 1, 7).setValues([
      [
        isEdit ? f.id : "TRX-" + Date.now(),
        tgl,
        nama,
        jenis,
        kategori,
        nominal,
        String(f.keterangan || ""),
      ],
    ]);
    ensureKategori_(kategori, jenis);
    return ok_(
      isEdit ? "Transaksi berhasil diperbarui" : "Transaksi berhasil disimpan",
    );
  });
}

/**
 * Simpan banyak transaksi baru sekaligus (dipanggil dari form "tambah banyak data" di frontend).
 * Semua baris ditulis dalam satu batch + satu lock, jadi jauh lebih cepat & aman dibanding
 * memanggil simpanDataTransaksi() berkali-kali dari client.
 */
function simpanBanyakTransaksi(list) {
  return tx_(() => {
    if (!Array.isArray(list) || !list.length)
      return err_("Tidak ada data yang dikirim");
    if (list.length > 50) return err_("Maksimal 50 data sekaligus");

    const rows = [];
    const now = Date.now();
    for (let i = 0; i < list.length; i++) {
      const f = list[i] || {};
      const nama = String(f.nama || "").trim(),
        nominal = Number(f.nominal);
      if (!nama || !(nominal > 0))
        return err_("Data " + (i + 1) + ": nama dan nominal wajib diisi");
      const jenis = f.jenis === "Pemasukan" ? "Pemasukan" : "Pengeluaran";
      const kategori = String(f.kategori || "").trim() || "Umum";
      const tgl = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(f.tanggal || "")
        ? f.tanggal
        : Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "yyyy-MM-dd HH:mm:ss",
          );
      // Sertakan index agar ID tetap unik walau disimpan pada milidetik yang sama.
      rows.push([
        "TRX-" + now + "-" + i,
        tgl,
        nama,
        jenis,
        kategori,
        nominal,
        String(f.keterangan || ""),
      ]);
    }

    const s = getSheet_(SH_TRX),
      start = s.getLastRow() + 1;
    s.getRange(start, 2, rows.length, 1).setNumberFormat("@"); // kolom tanggal sebagai teks
    s.getRange(start, 1, rows.length, 7).setValues(rows);

    // Pastikan semua kategori baru (unik per jenis) terdaftar di sheet Kategori.
    const uniqKat = {};
    rows.forEach((r) => {
      uniqKat[r[4] + "|" + r[3]] = true;
    });
    Object.keys(uniqKat).forEach((k) => {
      const idx = k.lastIndexOf("|");
      ensureKategori_(k.slice(0, idx), k.slice(idx + 1));
    });

    return ok_(
      rows.length > 1
        ? rows.length + " transaksi berhasil disimpan"
        : "Transaksi berhasil disimpan",
    );
  });
}

function hapusTransaksi(id) {
  return tx_(() => {
    const s = getSheet_(SH_TRX),
      row = findBy_(s, 1, id);
    if (!row) return err_("Data tidak ditemukan");
    s.deleteRow(row);
    return ok_("Transaksi berhasil dihapus");
  });
}

/**
 * Hapus banyak transaksi sekaligus berdasarkan daftar ID (dipanggil dari mode
 * "pilih banyak" di frontend). Dihapus dari baris paling bawah ke atas supaya
 * nomor baris yang belum diproses tidak ikut geser.
 */
function hapusTransaksiBanyak(ids) {
  return tx_(() => {
    if (!Array.isArray(ids) || !ids.length)
      return err_("Tidak ada data yang dipilih");
    const s = getSheet_(SH_TRX),
      last = s.getLastRow();
    if (last < 2) return err_("Data tidak ditemukan");
    const idSet = new Set(ids.map(String));
    const kolomId = s.getRange(2, 1, last - 1, 1).getValues();
    let jml = 0;
    for (let i = kolomId.length - 1; i >= 0; i--) {
      if (idSet.has(String(kolomId[i][0]))) {
        s.deleteRow(i + 2);
        jml++;
      }
    }
    if (!jml) return err_("Data tidak ditemukan, mungkin sudah dihapus");
    return ok_(jml + " transaksi berhasil dihapus");
  });
}

/**
 * Ubah jenis dan/atau kategori untuk banyak transaksi sekaligus (mode "pilih
 * banyak" di frontend). Field yang dikosongkan di form tidak akan diubah;
 * nama, nominal, dan tanggal per transaksi tetap seperti semula karena
 * memang khas per data dan tidak masuk akal diseragamkan.
 */
function editTransaksiBanyak(ids, patch) {
  return tx_(() => {
    if (!Array.isArray(ids) || !ids.length)
      return err_("Tidak ada data yang dipilih");
    patch = patch || {};
    const jenisBaru =
      patch.jenis === "Pemasukan" || patch.jenis === "Pengeluaran"
        ? patch.jenis
        : "";
    const katBaru = String(patch.kategori || "").trim();
    if (!jenisBaru && !katBaru) return err_("Tidak ada perubahan yang dipilih");

    const s = getSheet_(SH_TRX),
      last = s.getLastRow();
    if (last < 2) return err_("Data tidak ditemukan");
    const idSet = new Set(ids.map(String));
    const range = s.getRange(2, 1, last - 1, 7),
      data = range.getValues();
    let jml = 0;
    for (let i = 0; i < data.length; i++) {
      if (idSet.has(String(data[i][0]))) {
        if (jenisBaru) data[i][3] = jenisBaru;
        if (katBaru) data[i][4] = katBaru;
        jml++;
      }
    }
    if (!jml) return err_("Data tidak ditemukan, mungkin sudah dihapus");
    range.setValues(data);
    if (katBaru) ensureKategori_(katBaru, jenisBaru || "Pengeluaran");
    return ok_(jml + " transaksi berhasil diperbarui");
  });
}

/* ---------- Kategori ---------- */
function ensureKategori_(nama, tipe) {
  const k = getSheet_(SH_KAT);
  if (findBy_(k, 2, nama)) return;
  k.appendRow([Math.max(0, k.getLastRow() - 1) + 1, nama, tipe]);
}
function simpanKategori(nama, tipe) {
  return tx_(() => {
    nama = String(nama || "").trim();
    if (!nama) return err_("Nama kategori wajib diisi");
    if (findBy_(getSheet_(SH_KAT), 2, nama))
      return err_('Kategori "' + nama + '" sudah ada');
    ensureKategori_(nama, tipe === "Pemasukan" ? "Pemasukan" : "Pengeluaran");
    return ok_("Kategori berhasil ditambahkan");
  });
}
/**
 * Hapus kategori beserta SEMUA data yang terhubung dengannya:
 * - Semua baris Transaksi dengan kategori ini
 * - Baris Anggaran (batas per kategori) untuk kategori ini
 * Ini disengaja (cascade delete) supaya tidak ada data "yatim" yang nyantol
 * ke kategori yang sudah tidak ada.
 */
function hapusKategori(nama) {
  return tx_(() => {
    nama = String(nama || "").trim();
    if (!nama) return err_("Nama kategori wajib diisi");
    const sKat = getSheet_(SH_KAT),
      rowKat = findBy_(sKat, 2, nama);
    if (!rowKat) return err_("Kategori tidak ditemukan");
    const target = nama.toLowerCase();

    // 1) Hapus semua transaksi dengan kategori ini (dari bawah ke atas biar index tidak geser)
    const sTrx = getSheet_(SH_TRX),
      lastTrx = sTrx.getLastRow();
    let jmlTrx = 0;
    if (lastTrx > 1) {
      const dataTrx = sTrx.getRange(2, 5, lastTrx - 1, 1).getValues(); // kolom 5 = kategori
      for (let i = dataTrx.length - 1; i >= 0; i--) {
        if (String(dataTrx[i][0]).toLowerCase() === target) {
          sTrx.deleteRow(i + 2);
          jmlTrx++;
        }
      }
    }

    // 2) Hapus anggaran (batas bulanan) untuk kategori ini, kalau ada
    const sAng = getSheet_(SH_ANG),
      rowAng = findBy_(sAng, 1, nama);
    const adaAnggaran = !!rowAng;
    if (rowAng) sAng.deleteRow(rowAng);

    // 3) Hapus kategorinya sendiri (baris ini belum berubah karena penghapusan di atas ada di sheet lain)
    sKat.deleteRow(rowKat);

    const rincian = [];
    if (jmlTrx) rincian.push(jmlTrx + " transaksi");
    if (adaAnggaran) rincian.push("anggaran kategori ini");
    const pesan = rincian.length
      ? "Kategori dihapus, beserta " + rincian.join(" dan ")
      : "Kategori dihapus";
    return ok_(pesan);
  });
}

/* ---------- Pengaturan ---------- */
function simpanPengaturan(o) {
  return tx_(() => {
    const s = getSheet_(SH_SET);
    for (const k in DEF_SET_) {
      if (!(k in o)) continue;
      const n = Number(o[k]);
      if (isNaN(n) || n < 0) return err_('Nilai "' + k + '" tidak valid');
      const row = findBy_(s, 1, k);
      if (row) s.getRange(row, 2).setValue(n);
      else s.appendRow([k, n]);
    }
    return ok_("Pengaturan berhasil disimpan");
  });
}

/* ---------- Anggaran per kategori ---------- */
function simpanAnggaran(kategori, batas) {
  return tx_(() => {
    kategori = String(kategori || "").trim();
    batas = Number(batas);
    if (!kategori || !(batas > 0))
      return err_("Kategori dan batas anggaran wajib diisi");
    const s = getSheet_(SH_ANG),
      row = findBy_(s, 1, kategori);
    if (row) s.getRange(row, 2).setValue(batas);
    else s.appendRow([kategori, batas]);
    ensureKategori_(kategori, "Pengeluaran"); // pastikan kategori juga terdaftar di sheet Kategori
    return ok_("Anggaran berhasil disimpan");
  });
}
function hapusAnggaran(kategori) {
  return tx_(() => {
    const s = getSheet_(SH_ANG),
      row = findBy_(s, 1, kategori);
    if (!row) return err_("Anggaran tidak ditemukan");
    s.deleteRow(row);
    return ok_("Anggaran dihapus");
  });
}

/* ---------- Target tabungan ---------- */
function simpanTarget(o) {
  return tx_(() => {
    const nama = String(o.nama || "").trim(),
      target = Number(o.target),
      terkumpul = Number(o.terkumpul) || 0;
    if (!nama || !(target > 0) || terkumpul < 0)
      return err_("Nama dan nominal target wajib diisi");
    const tenggat = /^\d{4}-\d{2}-\d{2}$/.test(o.tenggat || "")
      ? o.tenggat
      : "";
    const s = getSheet_(SH_TGT);
    let row = o.id ? findBy_(s, 1, o.id) : 0;
    const isEdit = row > 0;
    if (!isEdit) row = s.getLastRow() + 1;
    s.getRange(row, 5).setNumberFormat("@");
    s.getRange(row, 1, 1, 5).setValues([
      [isEdit ? o.id : "TGT-" + Date.now(), nama, target, terkumpul, tenggat],
    ]);
    return ok_(
      isEdit ? "Target berhasil diperbarui" : "Target berhasil dibuat",
    );
  });
}
function hapusTarget(id) {
  return tx_(() => {
    const s = getSheet_(SH_TGT),
      row = findBy_(s, 1, id);
    if (!row) return err_("Target tidak ditemukan");
    s.deleteRow(row);
    return ok_("Target dihapus");
  });
}
function setorTarget(id, jumlah) {
  return tx_(() => {
    jumlah = Number(jumlah);
    if (!(jumlah > 0)) return err_("Jumlah setoran harus lebih dari 0");
    const s = getSheet_(SH_TGT),
      row = findBy_(s, 1, id);
    if (!row) return err_("Target tidak ditemukan");
    const target = Number(s.getRange(row, 3).getValue()) || 0;
    const baru = (Number(s.getRange(row, 4).getValue()) || 0) + jumlah;
    s.getRange(row, 4).setValue(baru);
    return ok_(
      baru >= target
        ? "Selamat, target tabungan tercapai!"
        : "Setoran berhasil ditambahkan",
    );
  });
}

/* ---------- Input suara ---------- */
/** Contoh ucapan: "beli makan siang 25 ribu", "terima gaji 5 juta", "bensin 20.000". */
function analisisSuaraPintar(teks) {
  try {
    const t = String(teks || "").toLowerCase();
    const m = t.match(
      /(\d{1,3}(?:\.\d{3})+|\d+(?:,\d+)?)\s*(ribu|rb|k|juta|jt)?/,
    );
    if (!m) return err_("Nominal tidak terdeteksi, coba sebutkan angkanya");
    let n = /\.\d{3}/.test(m[1])
      ? parseFloat(m[1].replace(/\./g, ""))
      : parseFloat(m[1].replace(",", "."));
    if (/^(ribu|rb|k)$/.test(m[2] || "")) n *= 1000;
    else if (/^(juta|jt)$/.test(m[2] || "")) n *= 1000000;
    const masuk = /terima|gaji|masuk|bonus|dapat|untung/.test(t);
    const kat = masuk
      ? /gaji/.test(t)
        ? "Gaji"
        : "Pemasukan Lain"
      : /makan|minum|kopi|jajan/.test(t)
        ? "Makanan"
        : /bensin|ojek|grab|gojek|parkir|tol|bus|kereta/.test(t)
          ? "Transportasi"
          : "Lainnya";
    return simpanDataTransaksi({
      nama: teks.charAt(0).toUpperCase() + teks.slice(1),
      jenis: masuk ? "Pemasukan" : "Pengeluaran",
      kategori: kat,
      nominal: Math.round(n),
      keterangan: "Dicatat via suara",
    });
  } catch (e) {
    return err_("Gagal memproses suara: " + e.message);
  }
}
