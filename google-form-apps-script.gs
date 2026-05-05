/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║   RoomBook — Google Apps Script                                  ║
 * ║   Menghubungkan Google Form ke RoomBook API                      ║
 * ╠══════════════════════════════════════════════════════════════════╣
 * ║  CARA PASANG:                                                    ║
 * ║  1. Buka Google Form Anda                                        ║
 * ║  2. Klik ⋮ (titik tiga) → "Script editor"                       ║
 * ║  3. Hapus kode default, paste seluruh isi file ini               ║
 * ║  4. Ganti CONFIG di bawah sesuai setup Anda                      ║
 * ║  5. Klik ikon jam (Triggers) → Add Trigger:                      ║
 * ║       Function: onFormSubmit                                     ║
 * ║       Event source: From form                                    ║
 * ║       Event type: On form submit                                 ║
 * ║  6. Authorize & Save                                             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ─── KONFIGURASI — SESUAIKAN INI ─────────────────────────────────────────────
const CONFIG = {
  // URL API RoomBook Anda (ganti dengan URL server yang sudah deploy)
  // Untuk development lokal, gunakan ngrok: https://xxx.ngrok.io
  API_BASE_URL: "https://your-roombook-api.com",

  // API Key untuk keamanan (harus sama dengan FORM_API_KEY di .env server)
  API_KEY: "roombook-form-key-2024",

  // Email admin yang menerima notifikasi jika booking GAGAL / konflik
  ADMIN_EMAIL: "zulfikar.dwi@majoo.id",

  // Nama organisasi untuk subject email
  ORG_NAME: "Majoo",
};

// ─── MAPPING NAMA ROOM KE ID ──────────────────────────────────────────────────
const ROOM_MAP = {
  "Raja Ampat":     1,
  "Bromo":          2,
  "Tegallalang":    3,
  "Tanjung Tinggi": 4,
};

// ─── MAIN TRIGGER — dipanggil otomatis saat form disubmit ────────────────────
function onFormSubmit(e) {
  try {
    var response   = e.response;
    var answers    = response.getItemResponses();
    var respondent = response.getRespondentEmail(); // email @majoo.id otomatis

    // Parse semua jawaban ke objek
    var data = parseFormAnswers(answers, respondent);

    // Validasi dasar
    var validationError = validateData(data);
    if (validationError) {
      sendFailureEmail(data.email || respondent, data.namaLengkap, validationError);
      return;
    }

    // Kirim ke RoomBook API
    var result = postToRoomBookAPI(data);

    if (result.success) {
      sendConfirmationEmail(data, result.booking);
      Logger.log("✅ Booking berhasil: " + JSON.stringify(result.booking));
    } else {
      sendFailureEmail(data.email, data.namaLengkap, result.error, data);
      Logger.log("❌ Booking gagal: " + result.error);
    }

  } catch (err) {
    Logger.log("ERROR onFormSubmit: " + err.toString());
    sendErrorToAdmin(err.toString(), e);
  }
}

// ─── PARSE JAWABAN FORM ───────────────────────────────────────────────────────
function parseFormAnswers(answers, respondentEmail) {
  var data = { email: respondentEmail };

  answers.forEach(function(item) {
    var title = item.getItem().getTitle().toLowerCase();
    var value = item.getResponse();

    if (title.includes("nama lengkap"))         data.namaLengkap   = value;
    else if (title.includes("email"))           data.email         = value || respondentEmail;
    else if (title.includes("departemen"))      data.departemen    = value;
    else if (title.includes("judul") || title.includes("nama") && title.includes("meeting"))
                                                data.judulMeeting  = value;
    else if (title.includes("ruang"))           data.namaRoom      = value;
    else if (title.includes("tanggal"))         data.tanggal       = formatDate(value);
    else if (title.includes("mulai"))           data.jamMulai      = value;
    else if (title.includes("selesai"))         data.jamSelesai    = value;
    else if (title.includes("jumlah peserta"))  data.jumlahPeserta = value;
    else if (title.includes("nama peserta"))    data.namaPeserta   = value;
    else if (title.includes("fasilitas"))       data.fasilitas     = Array.isArray(value) ? value.join(", ") : value;
    else if (title.includes("catatan"))         data.catatan       = value;
  });

  // Resolusi room ID
  data.roomId = ROOM_MAP[data.namaRoom] || null;

  return data;
}

// ─── FORMAT TANGGAL ke YYYY-MM-DD ─────────────────────────────────────────────
function formatDate(dateValue) {
  if (!dateValue) return "";
  // Google Forms bisa kembalikan Date object atau string
  var d = new Date(dateValue);
  if (isNaN(d.getTime())) return dateValue; // kembalikan as-is jika bukan date valid
  var yyyy = d.getFullYear();
  var mm   = String(d.getMonth() + 1).padStart(2, "0");
  var dd   = String(d.getDate()).padStart(2, "0");
  return yyyy + "-" + mm + "-" + dd;
}

// ─── VALIDASI ─────────────────────────────────────────────────────────────────
function validateData(data) {
  if (!data.namaLengkap)  return "Nama lengkap tidak boleh kosong.";
  if (!data.email)        return "Email tidak ditemukan.";
  if (!data.judulMeeting) return "Judul meeting tidak boleh kosong.";
  if (!data.roomId)       return "Room '" + data.namaRoom + "' tidak ditemukan. Pilihan valid: Raja Ampat, Bromo, Tegallalang, Tanjung Tinggi.";
  if (!data.tanggal)      return "Tanggal meeting tidak boleh kosong.";
  if (!data.jamMulai)     return "Jam mulai tidak boleh kosong.";
  if (!data.jamSelesai)   return "Jam selesai tidak boleh kosong.";
  if (data.jamMulai >= data.jamSelesai) return "Jam selesai harus lebih besar dari jam mulai.";
  return null; // valid
}

// ─── POST KE ROOMBOOK API ─────────────────────────────────────────────────────
function postToRoomBookAPI(data) {
  var payload = {
    roomId:       data.roomId,
    title:        data.judulMeeting,
    date:         data.tanggal,
    startTime:    data.jamMulai,
    endTime:      data.jamSelesai,
    attendees:    data.namaPeserta   || "",
    notes:        buildNotes(data),
    bookedByName: data.namaLengkap,
    bookedByEmail: data.email,
    source:       "google-form",
  };

  var options = {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify(payload),
    headers: {
      "x-api-key": CONFIG.API_KEY,
    },
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch(CONFIG.API_BASE_URL + "/api/bookings/form-submit", options);
    var code     = response.getResponseCode();
    var body     = JSON.parse(response.getContentText());

    if (code === 201 || code === 200) {
      return { success: true, booking: body };
    } else if (code === 409) {
      // Konflik jadwal
      var conflict = body.conflict || {};
      return {
        success: false,
        error: "Jadwal KONFLIK dengan booking yang sudah ada:\n" +
               "• Meeting: " + (conflict.title || "-") + "\n" +
               "• Waktu: " + (conflict.startTime || "-") + " – " + (conflict.endTime || "-") + "\n" +
               "• Dibooking oleh: " + (conflict.bookedByName || "-") + "\n\n" +
               "Silakan pilih waktu atau ruangan yang berbeda."
      };
    } else {
      return { success: false, error: "API error " + code + ": " + (body.error || response.getContentText()) };
    }
  } catch (err) {
    return { success: false, error: "Tidak dapat terhubung ke RoomBook API: " + err.toString() };
  }
}

// ─── BUILD NOTES dari data form ───────────────────────────────────────────────
function buildNotes(data) {
  var parts = [];
  if (data.departemen)    parts.push("Dept: " + data.departemen);
  if (data.jumlahPeserta) parts.push("Peserta: " + data.jumlahPeserta);
  if (data.fasilitas)     parts.push("Fasilitas: " + data.fasilitas);
  if (data.catatan)       parts.push(data.catatan);
  return parts.join(" | ");
}

// ─── EMAIL KONFIRMASI BERHASIL ────────────────────────────────────────────────
function sendConfirmationEmail(data, booking) {
  var subject = "✅ [RoomBook] Booking Dikonfirmasi — " + data.judulMeeting;
  var body = `
Halo ${data.namaLengkap},

Booking meeting Anda telah berhasil dikonfirmasi! 🎉

━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 DETAIL BOOKING
━━━━━━━━━━━━━━━━━━━━━━━━━━━
Meeting    : ${data.judulMeeting}
Ruangan    : ${data.namaRoom}
Tanggal    : ${formatDateReadable(data.tanggal)}
Waktu      : ${data.jamMulai} – ${data.jamSelesai}
Peserta    : ${data.namaPeserta || "-"}
Departemen : ${data.departemen || "-"}
Catatan    : ${data.catatan || "-"}
━━━━━━━━━━━━━━━━━━━━━━━━━━━

Booking ID : #${booking.id || "-"}

Jika ada perubahan atau pembatalan, silakan hubungi admin atau login ke RoomBook.

Salam,
Tim ${CONFIG.ORG_NAME}
  `.trim();

  GmailApp.sendEmail(data.email, subject, body);
}

// ─── EMAIL GAGAL / KONFLIK ────────────────────────────────────────────────────
function sendFailureEmail(email, nama, errorMessage, data) {
  var subject = "⚠️ [RoomBook] Booking Gagal — " + (data ? data.judulMeeting : "Unknown");
  var body = `
Halo ${nama || ""},

Maaf, booking Anda TIDAK berhasil diproses.

Alasan: ${errorMessage}

${data ? `Detail yang Anda kirim:
• Meeting    : ${data.judulMeeting || "-"}
• Ruangan    : ${data.namaRoom || "-"}
• Tanggal    : ${data.tanggal || "-"}
• Waktu      : ${data.jamMulai || "-"} – ${data.jamSelesai || "-"}` : ""}

Silakan isi ulang form dengan jadwal yang berbeda, atau hubungi admin di ${CONFIG.ADMIN_EMAIL}.

Salam,
Tim ${CONFIG.ORG_NAME}
  `.trim();

  GmailApp.sendEmail(email, subject, body);

  // Notifikasi ke admin juga
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    "⚠️ [RoomBook] Booking gagal dari: " + (nama || email),
    "Booking gagal:\n\n" + errorMessage + "\n\nDari: " + (nama || "") + " <" + email + ">"
  );
}

// ─── NOTIFIKASI ERROR KE ADMIN ────────────────────────────────────────────────
function sendErrorToAdmin(errorStr, event) {
  GmailApp.sendEmail(
    CONFIG.ADMIN_EMAIL,
    "🔴 [RoomBook] Script Error",
    "Terjadi error di Apps Script:\n\n" + errorStr + "\n\nEvent: " + JSON.stringify(event || {})
  );
}

// ─── FORMAT TANGGAL ke bahasa Indonesia ──────────────────────────────────────
function formatDateReadable(dateStr) {
  if (!dateStr) return "-";
  var months = ["Januari","Februari","Maret","April","Mei","Juni",
                "Juli","Agustus","September","Oktober","November","Desember"];
  var days   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var d = new Date(dateStr + "T00:00:00");
  return days[d.getDay()] + ", " + d.getDate() + " " + months[d.getMonth()] + " " + d.getFullYear();
}

// ─── TEST MANUAL (jalankan dari editor untuk coba) ───────────────────────────
function testManual() {
  var fakeData = {
    namaLengkap:   "Zulfikar Dwi",
    email:         "zulfikar.dwi@majoo.id",
    departemen:    "Product",
    judulMeeting:  "Test Booking dari Apps Script",
    namaRoom:      "Raja Ampat",
    roomId:        1,
    tanggal:       "2026-05-10",
    jamMulai:      "10:00",
    jamSelesai:    "11:00",
    namaPeserta:   "Zulfikar, Firjayanti",
    jumlahPeserta: "2 orang",
    fasilitas:     "TV, Whiteboard",
    catatan:       "Test dari Apps Script",
  };

  Logger.log("Data yang akan dikirim: " + JSON.stringify(fakeData));
  var result = postToRoomBookAPI(fakeData);
  Logger.log("Hasil: " + JSON.stringify(result));
}
