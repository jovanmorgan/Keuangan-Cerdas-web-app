/**
 * KEUANGAN CERDAS - Backend (Google Apps Script, container-bound ke Spreadsheet)
 * Versi Terbaru: Google Identity Auth, Multi-User Data Isolation, Device Security & WebAuthn
 */

const SH_TRX = "Transaksi",
  SH_KAT = "Kategori",
  SH_SET = "Pengaturan",
  SH_ANG = "Anggaran",
  SH_TGT = "Target",
  SH_USERS = "Users",
  SH_SESSIONS = "UserSessions",
  SH_OTP = "OtpStore";

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

const DEFAULT_CATEGORIES = [
  { nama: "Gaji", tipe: "Pemasukan" },
  { nama: "Bonus", tipe: "Pemasukan" },
  { nama: "Investasi", tipe: "Pemasukan" },
  { nama: "Pemasukan Lain", tipe: "Pemasukan" },
  { nama: "Makanan & Minuman", tipe: "Pengeluaran" },
  { nama: "Transportasi", tipe: "Pengeluaran" },
  { nama: "Belanja", tipe: "Pengeluaran" },
  { nama: "Tagihan & Utilitas", tipe: "Pengeluaran" },
  { nama: "Hiburan", tipe: "Pengeluaran" },
  { nama: "Kesehatan", tipe: "Pengeluaran" },
  { nama: "Pendidikan", tipe: "Pengeluaran" },
  { nama: "Lainnya", tipe: "Pengeluaran" },
];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.fn) {
    let args = [];
    try {
      args = p.args ? JSON.parse(p.args) : [];
    } catch (_) {
      args = [];
    }
    return handleApi_(p.fn, args, p.token, p.userToken);
  }
  setupDatabase();
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Keuangan Cerdas")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    return handleApi_(body.fn, body.args, body.token, body.userToken);
  } catch (err) {
    return jsonOutput_(err_("Permintaan tidak valid: " + err.message));
  }
}

/**
 * Whitelist fungsi API publik
 */
const API_FUNCTIONS_ = {
  // Data user functions (memerlukan activeUserId)
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
  prefetchData: prefetchData,

  // Google & Auth functions
  googleAuthCheck: googleAuthCheck,
  googleRegisterWithPin: googleRegisterWithPin,
  verifyNewDeviceWithPin: verifyNewDeviceWithPin,
  registerUser: registerUser,
  loginWithPin: loginWithPin,
  loginWithBiometric: loginWithBiometric,
  checkDeviceBiometric: checkDeviceBiometric,
  registerBiometric: registerBiometric,
  validateSession: validateSession,
  logoutSession: logoutSession,
  sendOtpEmail: sendOtpEmail,
  verifyOtp: verifyOtp,
  resetPin: resetPin,
  ping: ping,
};

function ping() {
  return { status: "success", message: "pong", timestamp: new Date().toISOString() };
}

function checkToken_(token) {
  const required = PropertiesService.getScriptProperties().getProperty("APP_TOKEN");
  if (!required) return true;
  return String(token || "") === String(required);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

/**
 * Dispatcher API utama dengan penanganan sesi & isolasi multi-user
 */
function handleApi_(fn, args, token, userToken) {
  if (!checkToken_(token)) return jsonOutput_(err_("Token server tidak valid"));
  if (!fn || !API_FUNCTIONS_.hasOwnProperty(fn))
    return jsonOutput_(err_("Fungsi tidak dikenali: " + fn));

  try {
    setupDatabase();

    const AUTH_PUBLIC_FNS = [
      "ping",
      "googleAuthCheck",
      "googleRegisterWithPin",
      "verifyNewDeviceWithPin",
      "registerUser",
      "loginWithPin",
      "loginWithBiometric",
      "checkDeviceBiometric",
      "validateSession",
      "logoutSession",
      "sendOtpEmail",
      "verifyOtp",
      "resetPin",
    ];

    let argList = Array.isArray(args) ? args.slice() : [];

    // Jika fungsi data butuh sesi user:
    if (!AUTH_PUBLIC_FNS.includes(fn)) {
      // Ambil userToken dari parameter atau argumen terakhir
      let uTok = userToken || "";
      if (!uTok && argList.length > 0 && typeof argList[argList.length - 1] === "string" && argList[argList.length - 1].length === 64) {
        // Fallback jika dikirim sebagai parameter terakhir
        uTok = argList.pop();
      }
      const sess = validateSession_(uTok);
      if (!sess) {
        return jsonOutput_(err_("Sesi tidak valid atau telah berakhir. Silakan login kembali."));
      }
      // Suntikkan activeUserId ke fungsi data
      argList.push(sess.userId);
    }

    const result = API_FUNCTIONS_[fn].apply(null, argList);
    return jsonOutput_(result === undefined ? ok_("") : result);
  } catch (err) {
    return jsonOutput_(err_("Gagal memproses permintaan: " + err.message));
  }
}

/* ---------- Helper Database ---------- */
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

function fmtTanggal_(v) {
  return v instanceof Date
    ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    : String(v || "");
}

/**
 * Setup Database dengan skema Multi-User (kolom UserId)
 * Mendukung migrasi otomatis jika sheet sudah ada sebelumnya.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    [SH_TRX]: ["Id", "Tanggal", "Nama", "Jenis", "Kategori", "Nominal", "Keterangan", "UserId"],
    [SH_KAT]: ["Id", "Nama", "Tipe", "UserId"],
    [SH_SET]: ["UserId", "Key", "Value"],
    [SH_ANG]: ["Kategori", "Batas", "UserId"],
    [SH_TGT]: ["Id", "Nama", "Target", "Terkumpul", "Tenggat", "UserId"],
    [SH_USERS]: ["UserId", "Nama", "Email", "PinHash", "DeviceIds", "BiometricCreds", "Picture", "CreatedAt", "LastLogin", "GoogleSub"],
    [SH_SESSIONS]: ["Token", "UserId", "DeviceId", "CreatedAt", "ExpiresAt", "IsValid"],
    [SH_OTP]: ["Email", "Otp", "ExpiresAt", "Used"]
  };

  for (const [name, headers] of Object.entries(sheets)) {
    let s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      s.getRange(1, 1, 1, headers.length).setValues([headers]);
      s.setFrozenRows(1);
    } else {
      // Auto-migration: Cek jika kolom UserId belum ada di sheet data
      const lastCol = s.getLastColumn();
      if (lastCol > 0) {
        const curHeaders = s.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
        if (headers.includes("UserId") && !curHeaders.includes("UserId")) {
          s.getRange(1, lastCol + 1).setValue("UserId");
        }
      }
    }
  }

  // Isi kategori default sistem jika Kategori masih kosong
  const katSheet = ss.getSheetByName(SH_KAT);
  if (katSheet && katSheet.getLastRow() < 2) {
    const defaultRows = DEFAULT_CATEGORIES.map((c, i) => [i + 1, c.nama, c.tipe, "SYSTEM"]);
    katSheet.getRange(2, 1, defaultRows.length, 4).setValues(defaultRows);
  }
}

/* ============================================================
 * OPERASI DATA MULTI-USER (ISOLASI BERDASARKAN activeUserId)
 * ============================================================ */

function getSettings_(activeUserId) {
  const o = Object.assign({}, DEF_SET_);
  const s = getSheet_(SH_SET);
  const last = s.getLastRow();
  if (last < 2) return o;
  const data = s.getRange(2, 1, last - 1, 3).getValues();
  data.forEach((r) => {
    if (String(r[0]) === String(activeUserId) && r[1] in DEF_SET_ && r[2] !== "" && !isNaN(r[2])) {
      o[r[1]] = Number(r[2]);
    }
  });
  return o;
}

/**
 * Baca semua data keuangan milik activeUserId
 */
function getAllData(activeUserId) {
  if (!activeUserId) return err_("User tidak terotorisasi");

  // 1. Transaksi (kolom 8 = UserId)
  const sTrx = getSheet_(SH_TRX);
  const lastTrx = sTrx.getLastRow();
  let trx = [];
  if (lastTrx > 1) {
    const rows = sTrx.getRange(2, 1, lastTrx - 1, 8).getValues();
    trx = rows
      .filter((r) => String(r[7] || "") === String(activeUserId))
      .map((r) => ({
        id: String(r[0]),
        tanggal: fmtTanggal_(r[1]),
        nama: String(r[2] || "-"),
        jenis: r[3] === "Pemasukan" ? "Pemasukan" : "Pengeluaran",
        kategori: String(r[4] || "Umum"),
        nominal: Number(r[5]) || 0,
        keterangan: String(r[6] || ""),
      }));
    trx.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  }

  // 2. Kategori (kolom 4 = UserId. Ambil SYSTEM atau milik user)
  const sKat = getSheet_(SH_KAT);
  const lastKat = sKat.getLastRow();
  let kategori = [];
  if (lastKat > 1) {
    const rows = sKat.getRange(2, 1, lastKat - 1, 4).getValues();
    kategori = rows
      .filter((r) => {
        const u = String(r[3] || "");
        return !u || u === "SYSTEM" || u === String(activeUserId);
      })
      .map((r) => ({ id: r[0], nama: String(r[1]), tipe: String(r[2]) }))
      .filter((k) => k.nama);
  }

  // 3. Anggaran (kolom 3 = UserId)
  const sAng = getSheet_(SH_ANG);
  const lastAng = sAng.getLastRow();
  let anggaran = [];
  if (lastAng > 1) {
    const rows = sAng.getRange(2, 1, lastAng - 1, 3).getValues();
    anggaran = rows
      .filter((r) => String(r[2] || "") === String(activeUserId))
      .map((r) => ({
        kategori: String(r[0]),
        batas: Number(r[1]) || 0,
      }));
  }

  // 4. Target (kolom 6 = UserId)
  const sTgt = getSheet_(SH_TGT);
  const lastTgt = sTgt.getLastRow();
  let target = [];
  if (lastTgt > 1) {
    const rows = sTgt.getRange(2, 1, lastTgt - 1, 6).getValues();
    target = rows
      .filter((r) => String(r[5] || "") === String(activeUserId))
      .map((r) => ({
        id: String(r[0]),
        nama: String(r[1]),
        target: Number(r[2]) || 0,
        terkumpul: Number(r[3]) || 0,
        tenggat: fmtTanggal_(r[4]).slice(0, 10),
      }));
  }

  return {
    trx: trx,
    kategori: kategori,
    anggaran: anggaran,
    target: target,
    settings: getSettings_(activeUserId),
  };
}

function prefetchData(activeUserId) {
  return getAllData(activeUserId);
}

/* ---------- Transaksi ---------- */
function simpanDataTransaksi(f, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const nama = String(f.nama || "").trim(),
      nominal = Number(f.nominal);
    if (!nama || !(nominal > 0))
      return err_("Nama dan nominal (lebih dari 0) wajib diisi");
    const jenis = f.jenis === "Pemasukan" ? "Pemasukan" : "Pengeluaran";
    const kategori = String(f.kategori || "").trim() || "Umum";
    const tgl = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(f.tanggal || "")
      ? f.tanggal
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    const s = getSheet_(SH_TRX);
    const last = s.getLastRow();
    let row = 0;
    if (f.id && last > 1) {
      const ids = s.getRange(2, 1, last - 1, 8).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(f.id) && String(ids[i][7] || "") === String(activeUserId)) {
          row = i + 2;
          break;
        }
      }
      if (!row) return err_("Data tidak ditemukan atau bukan milik Anda");
    }

    const isEdit = row > 0;
    if (!isEdit) row = s.getLastRow() + 1;
    s.getRange(row, 2).setNumberFormat("@");
    s.getRange(row, 1, 1, 8).setValues([
      [
        isEdit ? f.id : "TRX-" + Date.now(),
        tgl,
        nama,
        jenis,
        kategori,
        nominal,
        String(f.keterangan || ""),
        activeUserId,
      ],
    ]);
    return ok_(isEdit ? "Transaksi berhasil diperbarui" : "Transaksi berhasil disimpan");
  });
}

function simpanBanyakTransaksi(list, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    if (!Array.isArray(list) || !list.length) return err_("Tidak ada data yang dikirim");
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
        : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

      rows.push([
        "TRX-" + now + "-" + i,
        tgl,
        nama,
        jenis,
        kategori,
        nominal,
        String(f.keterangan || ""),
        activeUserId,
      ]);
    }

    const s = getSheet_(SH_TRX);
    const start = s.getLastRow() + 1;
    s.getRange(start, 2, rows.length, 1).setNumberFormat("@");
    s.getRange(start, 1, rows.length, 8).setValues(rows);

    return ok_(rows.length + " transaksi berhasil disimpan");
  });
}

function hapusTransaksi(id, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const s = getSheet_(SH_TRX);
    const last = s.getLastRow();
    if (last < 2) return err_("Data tidak ditemukan");
    const data = s.getRange(2, 1, last - 1, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id) && String(data[i][7] || "") === String(activeUserId)) {
        s.deleteRow(i + 2);
        return ok_("Transaksi berhasil dihapus");
      }
    }
    return err_("Data tidak ditemukan atau bukan milik Anda");
  });
}

function hapusTransaksiBanyak(ids, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    if (!Array.isArray(ids) || !ids.length) return err_("Tidak ada data yang dipilih");
    const s = getSheet_(SH_TRX);
    const last = s.getLastRow();
    if (last < 2) return err_("Data tidak ditemukan");
    const idSet = new Set(ids.map(String));
    const data = s.getRange(2, 1, last - 1, 8).getValues();
    let jml = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      if (idSet.has(String(data[i][0])) && String(data[i][7] || "") === String(activeUserId)) {
        s.deleteRow(i + 2);
        jml++;
      }
    }
    if (!jml) return err_("Data tidak ditemukan atau bukan milik Anda");
    return ok_(jml + " transaksi berhasil dihapus");
  });
}

function editTransaksiBanyak(ids, patch, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    if (!Array.isArray(ids) || !ids.length) return err_("Tidak ada data yang dipilih");
    patch = patch || {};
    const jenisBaru = patch.jenis === "Pemasukan" || patch.jenis === "Pengeluaran" ? patch.jenis : "";
    const katBaru = String(patch.kategori || "").trim();
    if (!jenisBaru && !katBaru) return err_("Tidak ada perubahan yang dipilih");

    const s = getSheet_(SH_TRX);
    const last = s.getLastRow();
    if (last < 2) return err_("Data tidak ditemukan");
    const idSet = new Set(ids.map(String));
    const range = s.getRange(2, 1, last - 1, 8);
    const data = range.getValues();
    let jml = 0;
    for (let i = 0; i < data.length; i++) {
      if (idSet.has(String(data[i][0])) && String(data[i][7] || "") === String(activeUserId)) {
        if (jenisBaru) data[i][3] = jenisBaru;
        if (katBaru) data[i][4] = katBaru;
        jml++;
      }
    }
    if (!jml) return err_("Data tidak ditemukan atau bukan milik Anda");
    range.setValues(data);
    return ok_(jml + " transaksi berhasil diperbarui");
  });
}

/* ---------- Kategori ---------- */
function simpanKategori(nama, tipe, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    nama = String(nama || "").trim();
    if (!nama) return err_("Nama kategori wajib diisi");
    const s = getSheet_(SH_KAT);
    const last = s.getLastRow();
    if (last > 1) {
      const data = s.getRange(2, 1, last - 1, 4).getValues();
      const exists = data.some(
        (r) =>
          String(r[1]).toLowerCase() === nama.toLowerCase() &&
          (!r[3] || r[3] === "SYSTEM" || String(r[3]) === String(activeUserId)),
      );
      if (exists) return err_('Kategori "' + nama + '" sudah ada');
    }
    s.appendRow([Math.max(0, s.getLastRow() - 1) + 1, nama, tipe === "Pemasukan" ? "Pemasukan" : "Pengeluaran", activeUserId]);
    return ok_("Kategori berhasil ditambahkan");
  });
}

function hapusKategori(nama, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    nama = String(nama || "").trim();
    if (!nama) return err_("Nama kategori wajib diisi");
    const sKat = getSheet_(SH_KAT);
    const lastKat = sKat.getLastRow();
    if (lastKat < 2) return err_("Kategori tidak ditemukan");
    const dataKat = sKat.getRange(2, 1, lastKat - 1, 4).getValues();
    let rowKat = 0;
    for (let i = 0; i < dataKat.length; i++) {
      if (String(dataKat[i][1]).toLowerCase() === nama.toLowerCase() && String(dataKat[i][3]) === String(activeUserId)) {
        rowKat = i + 2;
        break;
      }
    }
    if (!rowKat) return err_("Kategori kustom milik Anda tidak ditemukan (kategori sistem bawaan tidak dapat dihapus)");

    // Hapus transaksi user dengan kategori ini
    const sTrx = getSheet_(SH_TRX);
    const lastTrx = sTrx.getLastRow();
    if (lastTrx > 1) {
      const dataTrx = sTrx.getRange(2, 1, lastTrx - 1, 8).getValues();
      for (let i = dataTrx.length - 1; i >= 0; i--) {
        if (String(dataTrx[i][4]).toLowerCase() === nama.toLowerCase() && String(dataTrx[i][7] || "") === String(activeUserId)) {
          sTrx.deleteRow(i + 2);
        }
      }
    }

    // Hapus anggaran user untuk kategori ini
    const sAng = getSheet_(SH_ANG);
    const lastAng = sAng.getLastRow();
    if (lastAng > 1) {
      const dataAng = sAng.getRange(2, 1, lastAng - 1, 3).getValues();
      for (let i = dataAng.length - 1; i >= 0; i--) {
        if (String(dataAng[i][0]).toLowerCase() === nama.toLowerCase() && String(dataAng[i][2] || "") === String(activeUserId)) {
          sAng.deleteRow(i + 2);
        }
      }
    }

    sKat.deleteRow(rowKat);
    return ok_("Kategori berhasil dihapus");
  });
}

/* ---------- Pengaturan ---------- */
function simpanPengaturan(o, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const s = getSheet_(SH_SET);
    const last = s.getLastRow();
    const existing = {};
    if (last > 1) {
      const data = s.getRange(2, 1, last - 1, 3).getValues();
      data.forEach((r, idx) => {
        if (String(r[0]) === String(activeUserId)) existing[String(r[1])] = idx + 2;
      });
    }

    for (const k in DEF_SET_) {
      if (!(k in o)) continue;
      const n = Number(o[k]);
      if (isNaN(n) || n < 0) return err_('Nilai "' + k + '" tidak valid');
      if (existing[k]) {
        s.getRange(existing[k], 3).setValue(n);
      } else {
        s.appendRow([activeUserId, k, n]);
      }
    }
    return ok_("Pengaturan berhasil disimpan");
  });
}

/* ---------- Anggaran ---------- */
function simpanAnggaran(kategori, batas, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    kategori = String(kategori || "").trim();
    batas = Number(batas);
    if (!kategori || !(batas > 0)) return err_("Kategori dan batas anggaran wajib diisi");

    const s = getSheet_(SH_ANG);
    const last = s.getLastRow();
    let row = 0;
    if (last > 1) {
      const data = s.getRange(2, 1, last - 1, 3).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === kategori.toLowerCase() && String(data[i][2] || "") === String(activeUserId)) {
          row = i + 2;
          break;
        }
      }
    }

    if (row) s.getRange(row, 2).setValue(batas);
    else s.appendRow([kategori, batas, activeUserId]);
    return ok_("Anggaran berhasil disimpan");
  });
}

function hapusAnggaran(kategori, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const s = getSheet_(SH_ANG);
    const last = s.getLastRow();
    if (last < 2) return err_("Anggaran tidak ditemukan");
    const data = s.getRange(2, 1, last - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(kategori).toLowerCase() && String(data[i][2] || "") === String(activeUserId)) {
        s.deleteRow(i + 2);
        return ok_("Anggaran dihapus");
      }
    }
    return err_("Anggaran tidak ditemukan");
  });
}

/* ---------- Target Tabungan ---------- */
function simpanTarget(o, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const nama = String(o.nama || "").trim(),
      target = Number(o.target),
      terkumpul = Number(o.terkumpul) || 0;
    if (!nama || !(target > 0) || terkumpul < 0) return err_("Nama dan nominal target wajib diisi");
    const tenggat = /^\d{4}-\d{2}-\d{2}$/.test(o.tenggat || "") ? o.tenggat : "";

    const s = getSheet_(SH_TGT);
    const last = s.getLastRow();
    let row = 0;
    if (o.id && last > 1) {
      const data = s.getRange(2, 1, last - 1, 6).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]) === String(o.id) && String(data[i][5] || "") === String(activeUserId)) {
          row = i + 2;
          break;
        }
      }
      if (!row) return err_("Target tidak ditemukan");
    }

    const isEdit = row > 0;
    if (!isEdit) row = s.getLastRow() + 1;
    s.getRange(row, 5).setNumberFormat("@");
    s.getRange(row, 1, 1, 6).setValues([
      [isEdit ? o.id : "TGT-" + Date.now(), nama, target, terkumpul, tenggat, activeUserId],
    ]);
    return ok_(isEdit ? "Target berhasil diperbarui" : "Target berhasil dibuat");
  });
}

function hapusTarget(id, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    const s = getSheet_(SH_TGT);
    const last = s.getLastRow();
    if (last < 2) return err_("Target tidak ditemukan");
    const data = s.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id) && String(data[i][5] || "") === String(activeUserId)) {
        s.deleteRow(i + 2);
        return ok_("Target dihapus");
      }
    }
    return err_("Target tidak ditemukan");
  });
}

function setorTarget(id, jumlah, activeUserId) {
  return tx_(() => {
    if (!activeUserId) return err_("User tidak terotorisasi");
    jumlah = Number(jumlah);
    if (!(jumlah > 0)) return err_("Jumlah setoran harus lebih dari 0");
    const s = getSheet_(SH_TGT);
    const last = s.getLastRow();
    if (last < 2) return err_("Target tidak ditemukan");
    const data = s.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id) && String(data[i][5] || "") === String(activeUserId)) {
        const target = Number(data[i][2]) || 0;
        const baru = (Number(data[i][3]) || 0) + jumlah;
        s.getRange(i + 2, 4).setValue(baru);
        return ok_(baru >= target ? "Selamat, target tabungan tercapai! 🎉" : "Setoran berhasil ditambahkan");
      }
    }
    return err_("Target tidak ditemukan");
  });
}

function analisisSuaraPintar(teks, activeUserId) {
  try {
    const t = String(teks || "").toLowerCase();
    const m = t.match(/(\d{1,3}(?:\.\d{3})+|\d+(?:,\d+)?)\s*(ribu|rb|k|juta|jt)?/);
    if (!m) return err_("Nominal tidak terdeteksi, coba sebutkan angkanya");
    let n = /\.\d{3}/.test(m[1]) ? parseFloat(m[1].replace(/\./g, "")) : parseFloat(m[1].replace(",", "."));
    if (/^(ribu|rb|k)$/.test(m[2] || "")) n *= 1000;
    else if (/^(juta|jt)$/.test(m[2] || "")) n *= 1000000;
    const masuk = /terima|gaji|masuk|bonus|dapat|untung/.test(t);
    const kat = masuk
      ? /gaji/.test(t)
        ? "Gaji"
        : "Pemasukan Lain"
      : /makan|minum|kopi|jajan/.test(t)
        ? "Makanan & Minuman"
        : /bensin|ojek|grab|gojek|parkir|tol|bus|kereta/.test(t)
          ? "Transportasi"
          : "Lainnya";
    return simpanDataTransaksi(
      {
        nama: teks.charAt(0).toUpperCase() + teks.slice(1),
        jenis: masuk ? "Pemasukan" : "Pengeluaran",
        kategori: kat,
        nominal: Math.round(n),
        keterangan: "Dicatat via suara",
      },
      activeUserId,
    );
  } catch (e) {
    return err_("Gagal memproses suara: " + e.message);
  }
}

/* ============================================================
 * AUTENTIKASI GOOGLE, DEVICE SECURITY & WEBAUTHN
 * ============================================================ */

function findUserByEmail_(email) {
  const s = getSheet_(SH_USERS);
  const last = s.getLastRow();
  if (last < 2) return null;
  const data = s.getRange(2, 1, last - 1, 10).getValues();
  const emailLow = String(email || "").toLowerCase().trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === emailLow) {
      return { row: i + 2, data: data[i] };
    }
  }
  return null;
}

function generateToken_() {
  const bytes = [];
  for (let i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256));
  return bytes.map((b) => b.toString(16).padStart(2, "0")).join("");
}

function createSession_(userId, deviceId) {
  const token = generateToken_();
  const now = new Date();
  const expiry = new Date(now.getTime() + 12 * 60 * 60 * 1000); // Sesi 12 jam
  const s = getSheet_(SH_SESSIONS);
  s.appendRow([token, userId, deviceId, now.toISOString(), expiry.toISOString(), true]);

  // Update LastLogin di Users sheet
  const uSheet = getSheet_(SH_USERS);
  const last = uSheet.getLastRow();
  if (last >= 2) {
    const data = uSheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        uSheet.getRange(i + 2, 9).setValue(now.toISOString());
        break;
      }
    }
  }
  return { token, expiresAt: expiry.toISOString() };
}

function validateSession_(token) {
  if (!token) return null;
  const s = getSheet_(SH_SESSIONS);
  const last = s.getLastRow();
  if (last < 2) return null;
  const data = s.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(token) && data[i][5] === true) {
      const expiry = new Date(data[i][4]);
      if (expiry > new Date()) {
        return { userId: String(data[i][1]), deviceId: String(data[i][2]) };
      }
    }
  }
  return null;
}

/**
 * 1. GOOGLE AUTH CHECK:
 * Dipanggil saat user memilih akun Google / One Tap di browser.
 * Cek apakah user baru atau sudah terdaftar, dan apakah perangkat ini terotorisasi.
 */
function googleAuthCheck(payload, deviceId) {
  return tx_(() => {
    payload = payload || {};
    const email = String(payload.email || "").trim().toLowerCase();
    const nama = String(payload.nama || payload.name || "").trim();
    const picture = String(payload.picture || "").trim();
    const googleSub = String(payload.googleSub || payload.sub || "").trim();
    deviceId = String(deviceId || "").trim();

    if (!email) return err_("Email Google tidak valid");

    const user = findUserByEmail_(email);
    if (!user) {
      // User baru -> butuh setup PIN 6-digit
      return {
        status: "success",
        isNewUser: true,
        email: email,
        nama: nama,
        picture: picture,
        googleSub: googleSub,
      };
    }

    // User lama -> cek apakah deviceId ini sudah terdaftar
    let devices = [];
    try {
      devices = JSON.parse(user.data[4] || "[]");
    } catch (_) {
      devices = [];
    }
    const isAuthorizedDevice = deviceId && devices.includes(deviceId);

    // Cek apakah ada biometrik tersimpan untuk device ini
    let creds = [];
    try {
      creds = JSON.parse(user.data[5] || "[]");
    } catch (_) {
      creds = [];
    }
    const hasBio = creds.some((c) => c.deviceId === deviceId);

    if (isAuthorizedDevice) {
      // Perangkat terdaftar -> buka sesi langsung atau minta PIN/Bio cepat
      return {
        status: "success",
        isNewUser: false,
        isAuthorizedDevice: true,
        hasBio: hasBio,
        userId: String(user.data[0]),
        nama: String(user.data[1]),
        email: String(user.data[2]),
        picture: String(user.data[6] || picture),
      };
    } else {
      // Perangkat baru/lain -> LARANG akses langsung, minta verifikasi PIN
      return {
        status: "success",
        isNewUser: false,
        isAuthorizedDevice: false,
        requiresPinChallenge: true,
        userId: String(user.data[0]),
        nama: String(user.data[1]),
        email: String(user.data[2]),
        picture: String(user.data[6] || picture),
        message: "Perangkat baru terdeteksi! Masukkan PIN 6-digit untuk mendaftarkan perangkat ini.",
      };
    }
  });
}

/**
 * 2. REGISTRASI AKUN GOOGLE BARU:
 * Menyimpan akun baru dengan Google payload, PIN hash, dan mendaftarkan deviceId pertama.
 */
function googleRegisterWithPin(googleData, pinHash, deviceId) {
  return tx_(() => {
    googleData = googleData || {};
    const email = String(googleData.email || "").trim().toLowerCase();
    const nama = String(googleData.nama || googleData.name || "").trim() || "Pengguna";
    const picture = String(googleData.picture || "").trim();
    const googleSub = String(googleData.googleSub || googleData.sub || "").trim();
    pinHash = String(pinHash || "").trim();
    deviceId = String(deviceId || "").trim();

    if (!email || !pinHash) return err_("Email dan PIN wajib diisi");
    if (findUserByEmail_(email)) return err_("Akun dengan email ini sudah terdaftar");

    const userId = "USR-" + Date.now();
    const now = new Date().toISOString();
    const devices = deviceId ? [deviceId] : [];

    const s = getSheet_(SH_USERS);
    s.appendRow([
      userId,
      nama,
      email,
      pinHash,
      JSON.stringify(devices),
      "[]",
      picture,
      now,
      now,
      googleSub,
    ]);

    const session = createSession_(userId, deviceId);
    return {
      status: "success",
      userId: userId,
      nama: nama,
      email: email,
      picture: picture,
      token: session.token,
      expiresAt: session.expiresAt,
    };
  });
}

/**
 * 3. VERIFIKASI PERANGKAT BARU DENGAN PIN:
 * Jika user login dari perangkat lain, sistem meminta PIN. Jika PIN benar,
 * perangkat tersebut langsung ditambahkan ke daftar DeviceIds user dan sesi dibuka.
 */
function verifyNewDeviceWithPin(email, pinHash, deviceId) {
  return tx_(() => {
    email = String(email || "").trim().toLowerCase();
    pinHash = String(pinHash || "").trim();
    deviceId = String(deviceId || "").trim();

    if (!email || !pinHash) return err_("Email dan PIN wajib diisi");
    const user = findUserByEmail_(email);
    if (!user) return err_("Akun tidak ditemukan");

    if (String(user.data[3]) !== pinHash) {
      return err_("PIN salah! Akses dari perangkat ini ditolak.");
    }

    // PIN Benar -> Daftarkan deviceId baru
    let devices = [];
    try {
      devices = JSON.parse(user.data[4] || "[]");
    } catch (_) {
      devices = [];
    }
    if (deviceId && !devices.includes(deviceId)) {
      devices.push(deviceId);
      getSheet_(SH_USERS).getRange(user.row, 5).setValue(JSON.stringify(devices));
    }

    const session = createSession_(String(user.data[0]), deviceId);
    return {
      status: "success",
      message: "Perangkat berhasil diverifikasi dan didaftarkan!",
      userId: String(user.data[0]),
      nama: String(user.data[1]),
      email: String(user.data[2]),
      picture: String(user.data[6] || ""),
      token: session.token,
      expiresAt: session.expiresAt,
    };
  });
}

/**
 * Login konvensional dengan PIN
 */
function loginWithPin(email, pinHash, deviceId) {
  return tx_(() => {
    email = String(email || "").trim().toLowerCase();
    pinHash = String(pinHash || "").trim();
    deviceId = String(deviceId || "").trim();

    if (!email || !pinHash) return err_("Email dan PIN wajib diisi");
    const user = findUserByEmail_(email);
    if (!user) return err_("Email tidak ditemukan");

    if (String(user.data[3]) !== pinHash) return err_("PIN salah");

    // Otomatis daftarkan device jika PIN benar
    let devices = [];
    try {
      devices = JSON.parse(user.data[4] || "[]");
    } catch (_) {
      devices = [];
    }
    if (deviceId && !devices.includes(deviceId)) {
      devices.push(deviceId);
      getSheet_(SH_USERS).getRange(user.row, 5).setValue(JSON.stringify(devices));
    }

    const session = createSession_(String(user.data[0]), deviceId);
    return {
      status: "success",
      userId: String(user.data[0]),
      nama: String(user.data[1]),
      email: String(user.data[2]),
      picture: String(user.data[6] || ""),
      token: session.token,
      expiresAt: session.expiresAt,
    };
  });
}

function registerUser(email, nama, pinHash, deviceId) {
  return googleRegisterWithPin({ email, nama, picture: "", sub: "" }, pinHash, deviceId);
}

/**
 * Cek ketersediaan biometrik untuk deviceId
 */
function checkDeviceBiometric(deviceId, credentialId) {
  deviceId = String(deviceId || "").trim();
  credentialId = String(credentialId || "").trim();
  if (!deviceId || !credentialId) return { status: "success", found: false };

  const s = getSheet_(SH_USERS);
  const last = s.getLastRow();
  if (last < 2) return { status: "success", found: false };
  const data = s.getRange(2, 1, last - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    let creds = [];
    try {
      creds = JSON.parse(data[i][5] || "[]");
    } catch (_) {
      creds = [];
    }
    const found = creds.find((c) => c.credentialId === credentialId && c.deviceId === deviceId);
    if (found) {
      return {
        status: "success",
        found: true,
        userId: String(data[i][0]),
        nama: String(data[i][1]),
        picture: String(data[i][6] || ""),
      };
    }
  }
  return { status: "success", found: false };
}

/**
 * Login via WebAuthn Biometrik
 */
function loginWithBiometric(credentialId, deviceId) {
  return tx_(() => {
    credentialId = String(credentialId || "").trim();
    deviceId = String(deviceId || "").trim();
    if (!credentialId || !deviceId) return err_("Credential biometrik tidak valid");

    const s = getSheet_(SH_USERS);
    const last = s.getLastRow();
    if (last < 2) return err_("Tidak ada pengguna terdaftar");
    const data = s.getRange(2, 1, last - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      let creds = [];
      try {
        creds = JSON.parse(data[i][5] || "[]");
      } catch (_) {
        creds = [];
      }
      const found = creds.find((c) => c.credentialId === credentialId && c.deviceId === deviceId);
      if (found) {
        const session = createSession_(String(data[i][0]), deviceId);
        return {
          status: "success",
          userId: String(data[i][0]),
          nama: String(data[i][1]),
          email: String(data[i][2]),
          picture: String(data[i][6] || ""),
          token: session.token,
          expiresAt: session.expiresAt,
        };
      }
    }
    return err_("Sidik jari tidak dikenali. Silakan login dengan PIN.");
  });
}

function registerBiometric(token, credentialId, deviceId) {
  return tx_(() => {
    const sess = validateSession_(token);
    if (!sess) return err_("Sesi tidak valid, silakan login ulang");
    credentialId = String(credentialId || "").trim();
    deviceId = String(deviceId || "").trim();
    if (!credentialId || !deviceId) return err_("Data biometrik tidak valid");

    const s = getSheet_(SH_USERS);
    const last = s.getLastRow();
    if (last < 2) return err_("User tidak ditemukan");
    const data = s.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(sess.userId)) {
        let creds = [];
        try {
          creds = JSON.parse(data[i][5] || "[]");
        } catch (_) {
          creds = [];
        }
        creds = creds.filter((c) => c.deviceId !== deviceId);
        creds.push({ credentialId, deviceId, registeredAt: new Date().toISOString() });
        s.getRange(i + 2, 6).setValue(JSON.stringify(creds));
        return ok_("Sidik jari berhasil didaftarkan");
      }
    }
    return err_("User tidak ditemukan");
  });
}

function validateSession(token) {
  const sess = validateSession_(token);
  if (!sess) return { status: "success", valid: false };
  const s = getSheet_(SH_USERS);
  const last = s.getLastRow();
  if (last < 2) return { status: "success", valid: false };
  const data = s.getRange(2, 1, last - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === sess.userId) {
      return {
        status: "success",
        valid: true,
        userId: sess.userId,
        nama: String(data[i][1]),
        email: String(data[i][2]),
        picture: String(data[i][6] || ""),
      };
    }
  }
  return { status: "success", valid: false };
}

function logoutSession(token) {
  return tx_(() => {
    if (!token) return ok_("Berhasil logout");
    const s = getSheet_(SH_SESSIONS);
    const last = s.getLastRow();
    if (last < 2) return ok_("Berhasil logout");
    const data = s.getRange(2, 1, last - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(token)) {
        s.getRange(i + 2, 6).setValue(false);
        break;
      }
    }
    return ok_("Berhasil logout");
  });
}

/* ============================================================
 * OTP & RESET PIN
 * ============================================================ */

function sendOtpEmail(email, context) {
  return tx_(() => {
    email = String(email || "").trim().toLowerCase();
    context = String(context || "register");
    if (!email || !/^[^@]+@[^@]+\.[^@]+$/.test(email)) return err_("Format email tidak valid");

    if (context === "reset") {
      const user = findUserByEmail_(email);
      if (!user) return err_("Email tidak terdaftar");
    }

    const otp = String(Math.floor(100000 + Math.random() * 900000));
    const now = new Date();
    const expiry = new Date(now.getTime() + 10 * 60 * 1000);

    const s = getSheet_(SH_OTP);
    const last = s.getLastRow();
    if (last >= 2) {
      const data = s.getRange(2, 1, last - 1, 4).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === email && !data[i][3]) {
          s.getRange(i + 2, 4).setValue(true);
        }
      }
    }

    s.appendRow([email, otp, expiry.toISOString(), false]);

    const subject = context === "reset" ? "Reset PIN - Keuangan Cerdas" : "Verifikasi Email - Keuangan Cerdas";
    const htmlBody = `<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;background:#f8fafc;padding:24px;border-radius:16px">
      <div style="background:linear-gradient(135deg,#6366f1,#8b5cf6);border-radius:12px;padding:20px;text-align:center;margin-bottom:20px">
        <h1 style="color:#fff;margin:0;font-size:24px">💰 Keuangan Cerdas</h1>
      </div>
      <h2 style="color:#1e293b;margin:0 0 8px">${context === "reset" ? "Reset PIN" : "Verifikasi Email"}</h2>
      <p style="color:#64748b;margin:0 0 24px">Masukkan kode OTP berikut di aplikasi:</p>
      <div style="background:#fff;border:2px solid #6366f1;border-radius:12px;padding:24px;text-align:center;margin-bottom:24px">
        <span style="font-size:36px;font-weight:900;letter-spacing:8px;color:#6366f1">${otp}</span>
      </div>
      <p style="color:#94a3b8;font-size:12px">Kode berlaku selama <strong>10 menit</strong>. Jangan bagikan kode ini ke siapapun.</p>
    </div>`;

    MailApp.sendEmail(email, subject, "Kode OTP Anda: " + otp, { htmlBody: htmlBody });
    return { status: "success", message: "OTP dikirim ke email Anda", expiresInMinutes: 10 };
  });
}

function verifyOtp(email, otp) {
  return tx_(() => {
    email = String(email || "").trim().toLowerCase();
    otp = String(otp || "").trim();
    if (!email || !otp) return err_("Email dan OTP wajib diisi");

    const s = getSheet_(SH_OTP);
    const last = s.getLastRow();
    if (last < 2) return err_("OTP tidak valid atau sudah kadaluarsa");
    const data = s.getRange(2, 1, last - 1, 4).getValues();
    let foundRow = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).toLowerCase().trim() === email && String(data[i][1]).trim() === otp && !data[i][3]) {
        const expiry = new Date(data[i][2]);
        if (expiry < new Date()) return err_("OTP sudah kadaluarsa. Minta OTP baru.");
        foundRow = i + 2;
        break;
      }
    }
    if (foundRow < 0) return err_("OTP salah. Coba lagi.");
    s.getRange(foundRow, 4).setValue(true);
    return { status: "success", message: "OTP valid", verified: true };
  });
}

function resetPin(email, otp, newPinHash) {
  return tx_(() => {
    email = String(email || "").trim().toLowerCase();
    otp = String(otp || "").trim();
    newPinHash = String(newPinHash || "").trim();
    if (!email || !otp || !newPinHash) return err_("Semua field wajib diisi");

    const otpResult = verifyOtp(email, otp);
    if (otpResult.status !== "success" || !otpResult.verified) {
      return err_(otpResult.message || "OTP tidak valid");
    }

    const user = findUserByEmail_(email);
    if (!user) return err_("User tidak ditemukan");
    getSheet_(SH_USERS).getRange(user.row, 4).setValue(newPinHash);

    // Invalidasi sesi user
    const sSheet = getSheet_(SH_SESSIONS);
    const last = sSheet.getLastRow();
    if (last >= 2) {
      const data = sSheet.getRange(2, 1, last - 1, 6).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][1]) === String(user.data[0])) {
          sSheet.getRange(i + 2, 6).setValue(false);
        }
      }
    }
    return { status: "success", message: "PIN berhasil direset. Silakan login dengan PIN baru." };
  });
}
