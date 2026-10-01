/**
 * Membuat struktur sheet otomatis + data contoh (hanya saat sheet belum ada).
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fresh = !ss.getSheetByName("Transaksi"); // data contoh hanya untuk instalasi baru

  let sT = ss.getSheetByName("Transaksi");
  if (!sT) {
    sT = ss.insertSheet("Transaksi");
    const h = [
      "ID",
      "Tanggal",
      "Nama_Transaksi",
      "Jenis",
      "Kategori",
      "Nominal",
      "Keterangan",
    ];
    sT.appendRow(h);
    styleHeader_(sT, h.length);
    sT.getRange("B:B").setNumberFormat("@"); // tanggal disimpan sebagai teks

    // [hari lalu, jam, nama, jenis, kategori, nominal, keterangan]
    const seed = [
      [
        0,
        "08:00:00",
        "Gaji Bulanan",
        "Pemasukan",
        "Gaji",
        5000000,
        "Gaji bulan ini",
      ],
      [
        0,
        "09:30:00",
        "Makan Siang",
        "Pengeluaran",
        "Makanan",
        25000,
        "Makan warung",
      ],
      [1, "18:10:00", "Bensin Motor", "Pengeluaran", "Transportasi", 30000, ""],
      [2, "12:20:00", "Makan Malam", "Pengeluaran", "Makanan", 45000, ""],
      [
        3,
        "10:00:00",
        "Proyek Freelance",
        "Pemasukan",
        "Freelance",
        750000,
        "Desain logo",
      ],
      [
        5,
        "19:45:00",
        "Listrik",
        "Pengeluaran",
        "Tagihan",
        250000,
        "Token listrik",
      ],
      [6, "08:15:00", "Ojek Online", "Pengeluaran", "Transportasi", 18000, ""],
      [8, "13:00:00", "Belanja Bulanan", "Pengeluaran", "Belanja", 350000, ""],
      [12, "20:00:00", "Nonton Bioskop", "Pengeluaran", "Hiburan", 60000, ""],
      [35, "08:00:00", "Gaji Bulan Lalu", "Pemasukan", "Gaji", 5000000, ""],
      [38, "15:30:00", "Belanja Bulanan", "Pengeluaran", "Belanja", 420000, ""],
      [42, "11:00:00", "Internet", "Pengeluaran", "Tagihan", 300000, ""],
    ];
    const tz = Session.getScriptTimeZone(),
      now = Date.now();
    const rows = seed.map((r, i) => [
      "TRX-" + (1001 + i),
      Utilities.formatDate(new Date(now - r[0] * 864e5), tz, "yyyy-MM-dd") +
        " " +
        r[1],
      r[2],
      r[3],
      r[4],
      r[5],
      r[6],
    ]);
    sT.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  let sK = ss.getSheetByName("Kategori");
  if (!sK) {
    sK = ss.insertSheet("Kategori");
    const h = ["ID_Kategori", "Nama_Kategori", "Tipe_Transaksi"];
    sK.appendRow(h);
    styleHeader_(sK, h.length);
    [
      ["Gaji", "Pemasukan"],
      ["Freelance", "Pemasukan"],
      ["Makanan", "Pengeluaran"],
      ["Transportasi", "Pengeluaran"],
      ["Tagihan", "Pengeluaran"],
      ["Belanja", "Pengeluaran"],
      ["Hiburan", "Pengeluaran"],
    ].forEach((r, i) => sK.appendRow([i + 1, r[0], r[1]]));
  }

  let sP = ss.getSheetByName("Pengaturan");
  if (!sP) {
    sP = ss.insertSheet("Pengaturan");
    const h = ["Parameter_Key", "Parameter_Value"];
    sP.appendRow(h);
    styleHeader_(sP, h.length);
    sP.appendRow(["MIN_SALDO_WARNING", 50000]);
  }

  let sA = ss.getSheetByName("Anggaran");
  if (!sA) {
    sA = ss.insertSheet("Anggaran");
    const h = ["Kategori", "Batas_Bulanan"];
    sA.appendRow(h);
    styleHeader_(sA, h.length);
    if (fresh)
      [
        ["Makanan", 800000],
        ["Transportasi", 300000],
        ["Belanja", 600000],
        ["Hiburan", 200000],
      ].forEach((r) => sA.appendRow(r));
  }

  let sG = ss.getSheetByName("Target");
  if (!sG) {
    sG = ss.insertSheet("Target");
    const h = ["ID", "Nama_Target", "Target_Nominal", "Terkumpul", "Tenggat"];
    sG.appendRow(h);
    styleHeader_(sG, h.length);
    sG.getRange("E:E").setNumberFormat("@");
    if (fresh) {
      const d = Utilities.formatDate(
        new Date(Date.now() + 180 * 864e5),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
      sG.appendRow(["TGT-1", "Dana Darurat", 10000000, 2500000, d]);
    }
  }
}

function styleHeader_(sheet, numCols) {
  const r = sheet.getRange(1, 1, 1, numCols);
  r.setBackground("#4F46E5").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);
}
