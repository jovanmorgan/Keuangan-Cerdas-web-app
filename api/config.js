// Vercel Serverless Function (auto-detected karena ada di folder /api).
// Membaca Environment Variables project dan mengembalikannya sebagai JSON,
// supaya frontend statis bisa "tersambung otomatis" tanpa URL ditulis di kode.
//
// Atur di Vercel: Project > Settings > Environment Variables
//   API_URL   = URL Web App Apps Script kamu (diakhiri /exec)
//   API_TOKEN = (opsional) token yang sama dengan APP_TOKEN di Script Properties
//
// CATATAN: nilai ini tetap terlihat oleh siapa pun yang membuka situs (lewat
// tab Network browser) — ini bukan cara menyembunyikan token dari pengguna
// situs, hanya supaya tidak perlu ditulis di kode/GitHub dan tidak perlu
// diisi manual di tiap browser. Keamanan sesungguhnya tetap dari APP_TOKEN
// di sisi Apps Script.
module.exports = (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    url: process.env.API_URL || "",
    token: process.env.API_TOKEN || "",
  });
};
