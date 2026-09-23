// Vercel Serverless Function (auto-detected karena ada di folder /api).
// Membaca Environment Variables project dan mengembalikannya sebagai JSON,
// supaya frontend statis bisa "tersambung otomatis" tanpa URL ditulis di kode.
//
// Atur di Vercel: Project > Settings > Environment Variables
//   API_URL   = URL Web App Apps Script kamu (diakhiri /exec)
//   API_TOKEN = (opsional) token yang sama dengan APP_TOKEN di Script Properties
//
const DEFAULT_URL = "https://script.google.com/macros/s/AKfycbyyGt0wZTShzENEV_ePptnSd1N3pNhdxjzV7CR-2Uu0q7Pe7ObUTN6Wa-YRESnaSnjh/exec";

module.exports = (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    url: process.env.API_URL || DEFAULT_URL,
    token: process.env.API_TOKEN || "",
  });
};
