/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║  RoomBook — Google Apps Script                                       ║
 * ║  Form → Email Konfirmasi + Sync ke Google Sheet (bridge ke HTML)     ║
 * ╠══════════════════════════════════════════════════════════════════════╣
 * ║  CARA PASANG:                                                        ║
 * ║  1. Buka Google Form Anda                                            ║
 * ║  2. Klik ⋮ (titik tiga) → "Script editor"                           ║
 * ║  3. Hapus semua kode default, paste seluruh isi file ini             ║
 * ║  4. Isi CONFIG di bawah (minimal EMAIL_ADMIN & SHEET_ID)             ║
 * ║  5. Klik Save (💾)                                                   ║
 * ║  6. Klik ikon Jam ⏰ (Triggers) → + Add Trigger:                    ║
 * ║       • Function to run : onFormSubmit                               ║
 * ║       • Event source    : From form                                  ║
 * ║       • Event type      : On form submit                             ║
 * ║  7. Klik Save → Authorize → Allow                                    ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 */

// ─── ⚙️  KONFIGURASI — WAJIB DIISI ──────────────────────────────────────────
var CONFIG = {

  // Email admin yang menerima notifikasi setiap ada booking baru
  EMAIL_ADMIN: "zulfikar.dwi@majoo.id",

  // Google Sheet sebagai jembatan ke HTML app
  SHEET_ID: "1QbM-K2hGJPE0J8jSi8NsOs2Hj5kt0A-3NY_1-LXNSqo",

  // Nama tab sheet — HARUS tab baru terpisah dari Google Form responses
  SHEET_TAB: "RoomBook Data",

  // Nama organisasi untuk template email
  ORG_NAME: "Majoo",

  // Warna header email (ungu RoomBook)
  EMAIL_COLOR: "#4f46e5",
};

// ─── MAPPING NAMA ROOM → ID & KAPASITAS ──────────────────────────────────────
var ROOM_DATA = {
  "Raja Ampat":     { id: 1, capacity: 7,  facilities: "TV, Whiteboard" },
  "Bromo":          { id: 2, capacity: 9,  facilities: "TV, Whiteboard, Camera Conference" },
  "Tegallalang":    { id: 3, capacity: 6,  facilities: "TV, Whiteboard, Camera Conference" },
  "Tanjung Tinggi": { id: 4, capacity: 6,  facilities: "TV, Whiteboard" },
};

// ─── NAMA HARI & BULAN INDONESIA ─────────────────────────────────────────────
var HARI   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
var BULAN  = ["Januari","Februari","Maret","April","Mei","Juni",
              "Juli","Agustus","September","Oktober","November","Desember"];

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN TRIGGER — otomatis dipanggil saat form disubmit
// ═══════════════════════════════════════════════════════════════════════════════
function onFormSubmit(e) {
  try {
    var resp      = e.response;
    var emailResp = resp.getRespondentEmail();
    var items     = resp.getItemResponses();
    var data      = parseAnswers(items, emailResp);

    Logger.log("Data parsed: " + JSON.stringify(data));

    // 1. Cek collision di Google Sheet
    Logger.log("Email pengisi: '" + data.email + "'");
    Logger.log("Tanggal: " + data.tanggal + " | Jam: " + data.jamMulai + "-" + data.jamSelesai + " | Room: " + data.namaRoom + " (id:" + data.roomId + ")");

    var conflict = cekCollision(data);
    Logger.log("Hasil cekCollision: " + (conflict ? JSON.stringify(conflict) : "null (tidak ada collision)"));

    if (conflict) {
      Logger.log("COLLISION dengan: " + JSON.stringify(conflict));
      // Kirim email gagal ke pengisi
      if (data.email) {
        sendEmailGagal(data, conflict);
        Logger.log("Email gagal terkirim ke: " + data.email);
      } else {
        Logger.log("PERINGATAN: data.email kosong, tidak bisa kirim email gagal ke pengisi.");
        // Kirim ke admin agar admin tahu ada yang gagal booking
        try {
          GmailApp.sendEmail(CONFIG.EMAIL_ADMIN,
            "[RoomBook] Booking Gagal - Email pengisi tidak tersedia",
            "Booking ditolak collision tapi email pengisi kosong.\n\nData: " + JSON.stringify(data) + "\n\nConflict: " + JSON.stringify(conflict));
        } catch(e3) { Logger.log("Gagal kirim email fallback: " + e3); }
      }
      // Kirim notifikasi collision ke admin
      sendEmailAdminCollision(data, conflict);
      Logger.log("Booking DITOLAK karena collision: " + data.judulMeeting);
      return; // stop — tidak simpan ke sheet
    }

    // 2. Tidak ada collision — proses normal
    if (data.email) sendEmailKonfirmasi(data);
    sendEmailAdmin(data);
    saveToSheet(data);

    Logger.log("Semua proses selesai untuk: " + data.judulMeeting);

  } catch (err) {
    Logger.log("ERROR: " + err.toString() + "\n" + err.stack);
    try {
      GmailApp.sendEmail(CONFIG.EMAIL_ADMIN,
        "[RoomBook] Error di Apps Script",
        "Terjadi error:\n\n" + err.toString() + "\n\nStack:\n" + err.stack);
    } catch(e2) {}
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// CEK COLLISION — bandingkan booking baru dengan data di sheet
// ═══════════════════════════════════════════════════════════════════════════════
function cekCollision(data) {
  if (!data.tanggal || !data.jamMulai || !data.jamSelesai || !data.roomId) return null;

  try {
    var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var tab = ss.getSheetByName(CONFIG.SHEET_TAB);
    if (!tab || tab.getLastRow() < 2) return null;

    var values = tab.getDataRange().getValues();
    var header = values[0];

    // Cari index kolom yang dibutuhkan
    var iRoom    = header.indexOf("roomId");
    var iTgl     = header.indexOf("tanggal");
    var iMulai   = header.indexOf("jamMulai");
    var iSelesai = header.indexOf("jamSelesai");
    var iJudul   = header.indexOf("judulMeeting");
    var iNama    = header.indexOf("namaLengkap");

    // Fallback: coba baca dari kolom format Google Form
    var iRoomName = header.indexOf("Pilih Ruang Meeting");
    var iTglForm  = header.indexOf("Tanggal Meeting");
    var iMulaiF   = header.indexOf("Waktu Mulai");
    var iSelesaiF = header.indexOf("Waktu Selesai");
    var iJudulF   = header.indexOf("Nama / Judul Meeting");
    var iNamaF    = header.indexOf("Nama Lengkap *");

    var newMulai   = toMinutes(data.jamMulai);
    var newSelesai = toMinutes(data.jamSelesai);

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var existRoomId, existTgl, existMulai, existSelesai, existJudul, existNama;

      // Format Apps Script
      if (iRoom >= 0 && row[iRoom]) {
        existRoomId  = parseInt(row[iRoom]);
        existTgl     = formatTanggalISO(row[iTgl]);
        existMulai   = formatJam(String(row[iMulai]));
        existSelesai = formatJam(String(row[iSelesai]));
        existJudul   = row[iJudul] || "";
        existNama    = row[iNama]  || "";
      }
      // Format Google Form
      else if (iRoomName >= 0 && row[iRoomName]) {
        var rn = String(row[iRoomName]);
        existRoomId  = ROOM_DATA[rn] ? ROOM_DATA[rn].id : 0;
        existTgl     = formatTanggalISO(row[iTglForm]);
        existMulai   = formatJam(String(row[iMulaiF]));
        existSelesai = formatJam(String(row[iSelesaiF]));
        existJudul   = row[iJudulF] || "";
        existNama    = iNamaF >= 0 ? (row[iNamaF] || "") : "";
      } else {
        continue;
      }

      if (!existRoomId || !existTgl || !existMulai || !existSelesai) continue;
      if (existRoomId !== data.roomId) continue;
      if (existTgl !== data.tanggal) continue;

      var eMulai   = toMinutes(existMulai);
      var eSelesai = toMinutes(existSelesai);

      // Cek overlap waktu
      if (newMulai < eSelesai && newSelesai > eMulai) {
        return { judulMeeting: existJudul, namaLengkap: existNama,
                 jamMulai: existMulai, jamSelesai: existSelesai };
      }
    }
    return null;
  } catch(err) {
    Logger.log("Error cekCollision: " + err.toString() + "\n" + err.stack);
    // Kirim alert ke admin jika cek collision gagal
    try {
      GmailApp.sendEmail(CONFIG.EMAIL_ADMIN,
        "[RoomBook] Error cekCollision",
        "Collision check gagal:\n\n" + err.toString() + "\n\nData booking:\n" + JSON.stringify(data));
    } catch(e2) {}
    return null; // jika gagal cek, tetap lanjutkan booking
  }
}

// Helper konversi jam ke menit — support: "HH:MM", "H:MM AM/PM", ISO "1899-12-30T02:00:00.000Z"
function toMinutes(jam) {
  if (!jam) return 0;
  var s = String(jam).trim();
  // ISO date string dari Google Sheets: "1899-12-30T02:00:00.000Z"
  // Jam UTC dalam string ini = jam lokal yang disimpan (karena epoch 1899)
  var iso = s.match(/T(\d{2}):(\d{2})/);
  if (iso) return parseInt(iso[1]) * 60 + parseInt(iso[2]);
  // AM/PM
  var ampm = s.match(/^(\d{1,2}):(\d{2})(?::\d+)?\s*(AM|PM)$/i);
  if (ampm) {
    var h = parseInt(ampm[1]), m = parseInt(ampm[2]);
    var p = ampm[3].toUpperCase();
    if (p === "AM" && h === 12) h = 0;
    if (p === "PM" && h !== 12) h += 12;
    return h * 60 + m;
  }
  // HH:MM atau H:MM
  var plain = s.match(/^(\d{1,2}):(\d{2})/);
  if (plain) return parseInt(plain[1]) * 60 + parseInt(plain[2]);
  return 0;
}

// ═══════════════════════════════════════════════════════════════════════════════
// PARSE JAWABAN FORM
// Sesuaikan bagian ini jika nama pertanyaan di form Anda berbeda
// ═══════════════════════════════════════════════════════════════════════════════
function parseAnswers(items, emailRespondent) {
  var data = {
    email:         emailRespondent || "",
    namaLengkap:   "",
    departemen:    "",
    judulMeeting:  "",
    namaRoom:      "",
    tanggal:       "",
    jamMulai:      "",
    jamSelesai:    "",
    jumlahPeserta: "",
    namaPeserta:   "",
    fasilitas:     "",
    catatan:       "",
    timestamp:     new Date().toISOString(),
  };

  items.forEach(function(item) {
    var pertanyaan = item.getItem().getTitle().toLowerCase().trim();
    var jawaban    = item.getResponse();

    // ── Identitas ──
    if (cek(pertanyaan, ["nama lengkap", "nama anda", "your name", "full name"])) {
      data.namaLengkap = jawaban;
    }
    else if (cek(pertanyaan, ["email", "alamat email"])) {
      data.email = jawaban || data.email;
    }
    else if (cek(pertanyaan, ["departemen", "divisi", "tim", "department", "team"])) {
      data.departemen = jawaban;
    }

    // ── Detail Meeting ──
    else if (cek(pertanyaan, ["judul meeting", "nama meeting", "meeting title", "agenda", "keperluan", "tujuan"])) {
      data.judulMeeting = jawaban;
    }
    else if (cek(pertanyaan, ["ruang", "room", "ruangan"])) {
      // Jawaban bisa berisi nama room lengkap, cari yang cocok
      data.namaRoom = extractRoomName(jawaban);
    }

    // ── Jadwal ──
    else if (cek(pertanyaan, ["tanggal", "date", "hari"])) {
      data.tanggal = formatTanggalISO(jawaban);
    }
    else if (cek(pertanyaan, ["jam mulai", "waktu mulai", "start time", "mulai", "dari jam"])) {
      data.jamMulai = formatJam(jawaban);
    }
    else if (cek(pertanyaan, ["jam selesai", "waktu selesai", "end time", "selesai", "sampai jam", "berakhir"])) {
      data.jamSelesai = formatJam(jawaban);
    }

    // ── Peserta ──
    else if (cek(pertanyaan, ["jumlah peserta", "jumlah orang", "berapa orang", "number of"])) {
      data.jumlahPeserta = jawaban;
    }
    else if (cek(pertanyaan, ["nama peserta", "peserta", "attendees", "participant"])) {
      data.namaPeserta = Array.isArray(jawaban) ? jawaban.join(", ") : jawaban;
    }
    else if (cek(pertanyaan, ["fasilitas", "facility", "peralatan", "kebutuhan"])) {
      data.fasilitas = Array.isArray(jawaban) ? jawaban.join(", ") : jawaban;
    }
    else if (cek(pertanyaan, ["catatan", "notes", "keterangan", "informasi tambahan", "lainnya"])) {
      data.catatan = jawaban;
    }
  });

  // Tambahkan info room
  var roomInfo = ROOM_DATA[data.namaRoom] || {};
  data.roomId       = roomInfo.id       || null;
  data.kapasitas    = roomInfo.capacity || "-";
  data.fasilitasRoom = roomInfo.facilities || "-";

  return data;
}

// ─── Helper: cek apakah string mengandung salah satu keyword ─────────────────
function cek(str, keywords) {
  return keywords.some(function(k) { return str.indexOf(k) !== -1; });
}

// ─── Helper: ekstrak nama room dari jawaban ───────────────────────────────────
function extractRoomName(jawaban) {
  var rooms = Object.keys(ROOM_DATA);
  for (var i = 0; i < rooms.length; i++) {
    if (jawaban.toLowerCase().indexOf(rooms[i].toLowerCase()) !== -1) {
      return rooms[i];
    }
  }
  return jawaban; // kembalikan apa adanya jika tidak cocok
}

// ─── Helper: format tanggal ke YYYY-MM-DD ────────────────────────────────────
function formatTanggalISO(val) {
  if (!val) return "";
  // Jika sudah format Date object dari Google Forms
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  // Jika string, coba parse
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

// ─── Helper: format jam ke HH:MM ─────────────────────────────────────────────
function formatJam(val) {
  if (!val) return "";
  // Jika Date object langsung
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  var str = String(val).trim();
  // Sudah format HH:MM
  if (/^\d{2}:\d{2}$/.test(str)) return str;
  // Format H:MM → 0H:MM
  if (/^\d{1}:\d{2}$/.test(str)) return "0" + str;
  // ISO string dari Google Sheets: "1899-12-30T02:00:00.000Z" → ambil jam UTC
  var isoMatch = str.match(/T(\d{2}):(\d{2})/);
  if (isoMatch) {
    // Jam di ISO string ini adalah UTC, tapi Google Sheets menyimpan waktu lokal
    // sebagai epoch 1899-12-30, jadi ambil langsung jam:menit dari string
    return isoMatch[1] + ":" + isoMatch[2];
  }
  // Format AM/PM: "4:00:00 PM"
  var ampm = str.match(/^(\d{1,2}):(\d{2})(?::\d+)?\s*(AM|PM)$/i);
  if (ampm) {
    var h = parseInt(ampm[1]), m = ampm[2];
    var p = ampm[3].toUpperCase();
    if (p === "AM" && h === 12) h = 0;
    if (p === "PM" && h !== 12) h += 12;
    return String(h).padStart(2,"0") + ":" + m;
  }
  return str;
}

// ─── Helper: format tanggal ke teks Indonesia ────────────────────────────────
function formatTanggalIndo(tgl) {
  if (!tgl) return "-";
  var d = new Date(tgl + "T00:00:00");
  if (isNaN(d.getTime())) return tgl;
  return HARI[d.getDay()] + ", " + d.getDate() + " " + BULAN[d.getMonth()] + " " + d.getFullYear();
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMAIL KONFIRMASI → ke pengisi form
// ═══════════════════════════════════════════════════════════════════════════════
function sendEmailKonfirmasi(data) {
  var subject = "[Booking Dikonfirmasi] " + (data.judulMeeting || "Meeting Anda") + " - Meeting Room Booking " + CONFIG.ORG_NAME;

  var pesertaRow   = data.namaPeserta ? "<tr><td style='padding:7px 0;color:#6b7280;width:140px;'>Peserta</td><td style='padding:7px 0;color:#111827;'>" + data.namaPeserta + "</td></tr>" : "";
  var departemenRow = data.departemen ? "<tr style='background:rgba(0,0,0,0.02);'><td style='padding:7px 8px;color:#6b7280;'>Departemen</td><td style='padding:7px 8px;color:#111827;'>" + data.departemen + "</td></tr>" : "";
  var catatanRow   = data.catatan     ? "<tr><td style='padding:7px 0;color:#6b7280;'>Catatan</td><td style='padding:7px 0;color:#111827;'>" + data.catatan + "</td></tr>" : "";

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8fafc;'>" +

    "<div style='background:" + CONFIG.EMAIL_COLOR + ";padding:28px 32px;border-radius:12px 12px 0 0;text-align:center;'>" +
      "<h1 style='color:white;margin:0;font-size:24px;font-weight:700;letter-spacing:-0.5px;'>Booking Dikonfirmasi</h1>" +
      "<p style='color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:14px;'>Meeting Room Booking &mdash; " + CONFIG.ORG_NAME + "</p>" +
    "</div>" +

    "<div style='background:white;padding:28px 32px;border-radius:0 0 12px 12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);'>" +

      "<p style='color:#374151;font-size:15px;margin:0 0 20px;'>" +
        "Halo <strong>" + (data.namaLengkap || "Karyawan") + "</strong>,<br>" +
        "Booking ruang meeting Anda telah <strong style='color:#059669;'>berhasil dikonfirmasi</strong>." +
      "</p>" +

      "<div style='background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:20px;margin-bottom:20px;'>" +
        "<h3 style='margin:0 0 14px;color:#065f46;font-size:15px;border-bottom:1px solid #bbf7d0;padding-bottom:10px;'>Detail Booking</h3>" +
        "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
          "<tr><td style='padding:7px 0;color:#6b7280;width:140px;'>Meeting</td><td style='padding:7px 0;color:#111827;font-weight:600;'>" + (data.judulMeeting || "-") + "</td></tr>" +
          "<tr style='background:rgba(0,0,0,0.02);'><td style='padding:7px 8px;color:#6b7280;'>Ruangan</td><td style='padding:7px 8px;color:#111827;font-weight:600;'>" + (data.namaRoom || "-") + "</td></tr>" +
          "<tr><td style='padding:7px 0;color:#6b7280;'>Tanggal</td><td style='padding:7px 0;color:#111827;'>" + formatTanggalIndo(data.tanggal) + "</td></tr>" +
          "<tr style='background:rgba(0,0,0,0.02);'><td style='padding:7px 8px;color:#6b7280;'>Waktu</td><td style='padding:7px 8px;color:#111827;font-weight:700;color:#4f46e5;'>" + (data.jamMulai || "-") + " &ndash; " + (data.jamSelesai || "-") + "</td></tr>" +
          pesertaRow + departemenRow + catatanRow +
        "</table>" +
      "</div>" +

      "<div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:16px;margin-bottom:20px;'>" +
        "<h3 style='margin:0 0 10px;color:#1e40af;font-size:14px;'>Info Ruangan: " + (data.namaRoom || "-") + "</h3>" +
        "<p style='margin:4px 0;font-size:13px;color:#374151;'>Kapasitas: <strong>" + data.kapasitas + " orang</strong></p>" +
        "<p style='margin:4px 0;font-size:13px;color:#374151;'>Fasilitas: <strong>" + data.fasilitasRoom + "</strong></p>" +
      "</div>" +

      "<div style='background:#fefce8;border:1px solid #fde68a;border-radius:10px;padding:14px;margin-bottom:20px;'>" +
        "<p style='margin:0;font-size:13px;color:#92400e;'>" +
          "<strong>Pengingat:</strong> Harap datang tepat waktu dan pastikan ruangan bersih setelah digunakan. " +
          "Jika perlu membatalkan, hubungi admin segera." +
        "</p>" +
      "</div>" +

      "<p style='font-size:13px;color:#6b7280;margin:0;'>" +
        "Email ini dikirim otomatis oleh sistem Meeting Room Booking " + CONFIG.ORG_NAME + ".<br>" +
        "Pertanyaan? Hubungi <a href='mailto:" + CONFIG.EMAIL_ADMIN + "' style='color:" + CONFIG.EMAIL_COLOR + ";'>" + CONFIG.EMAIL_ADMIN + "</a>" +
      "</p>" +

    "</div>" +
    "</div>";

  GmailApp.sendEmail(data.email, subject, stripHtml(html), { htmlBody: html, name: "Meeting Room Booking " + CONFIG.ORG_NAME });
  Logger.log("Email konfirmasi terkirim ke: " + data.email);
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMAIL NOTIFIKASI → ke admin
// ═══════════════════════════════════════════════════════════════════════════════
function sendEmailAdmin(data) {
  var subject = "[Booking Baru] " + (data.judulMeeting || "-") + " - " + (data.namaRoom || "-") + " - Meeting Room Booking";

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;'>" +
      "<div style='background:" + CONFIG.EMAIL_COLOR + ";padding:20px 24px;border-radius:10px 10px 0 0;'>" +
        "<h2 style='color:white;margin:0;font-size:18px;'>Booking Baru via Google Form</h2>" +
        "<p style='color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;'>Meeting Room Booking &mdash; " + CONFIG.ORG_NAME + "</p>" +
      "</div>" +
      "<div style='background:white;padding:24px;border:1px solid #e5e7eb;border-radius:0 0 10px 10px;'>" +
        "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
          "<tr style='background:#f9fafb;'><td style='padding:9px 12px;font-weight:600;color:#374151;width:150px;'>Meeting</td><td style='padding:9px 12px;'>" + (data.judulMeeting||"-") + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Ruangan</td><td style='padding:9px 12px;font-weight:700;color:" + CONFIG.EMAIL_COLOR + ";'>" + (data.namaRoom||"-") + "</td></tr>" +
          "<tr style='background:#f9fafb;'><td style='padding:9px 12px;font-weight:600;color:#374151;'>Tanggal</td><td style='padding:9px 12px;'>" + formatTanggalIndo(data.tanggal) + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Waktu</td><td style='padding:9px 12px;font-weight:700;color:" + CONFIG.EMAIL_COLOR + ";'>" + (data.jamMulai||"-") + " &ndash; " + (data.jamSelesai||"-") + "</td></tr>" +
          "<tr style='background:#f9fafb;'><td style='padding:9px 12px;font-weight:600;color:#374151;'>Dibooking oleh</td><td style='padding:9px 12px;'>" + (data.namaLengkap||"-") + " (" + (data.email||"-") + ")</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Departemen</td><td style='padding:9px 12px;'>" + (data.departemen||"-") + "</td></tr>" +
          "<tr style='background:#f9fafb;'><td style='padding:9px 12px;font-weight:600;color:#374151;'>Peserta</td><td style='padding:9px 12px;'>" + (data.namaPeserta||"-") + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Jumlah</td><td style='padding:9px 12px;'>" + (data.jumlahPeserta||"-") + "</td></tr>" +
          "<tr style='background:#f9fafb;'><td style='padding:9px 12px;font-weight:600;color:#374151;'>Catatan</td><td style='padding:9px 12px;'>" + (data.catatan||"-") + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Waktu Submit</td><td style='padding:9px 12px;color:#6b7280;font-size:12px;'>" + (data.timestamp||"-") + "</td></tr>" +
        "</table>" +
      "</div>" +
    "</div>";

  GmailApp.sendEmail(CONFIG.EMAIL_ADMIN, subject, stripHtml(html), { htmlBody: html, name: "Meeting Room Booking System" });
  Logger.log("Notifikasi admin terkirim ke: " + CONFIG.EMAIL_ADMIN);
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMAIL GAGAL BOOKING (COLLISION) → ke pengisi form
// ═══════════════════════════════════════════════════════════════════════════════
function sendEmailGagal(data, conflict) {
  var subject = "[Booking Gagal] " + (data.judulMeeting || "Meeting Anda") + " - Meeting Room Booking " + CONFIG.ORG_NAME;

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8fafc;'>" +

    "<div style='background:#dc2626;padding:28px 32px;border-radius:12px 12px 0 0;text-align:center;'>" +
      "<h1 style='color:white;margin:0;font-size:24px;font-weight:700;'>Booking Gagal</h1>" +
      "<p style='color:rgba(255,255,255,0.85);margin:8px 0 0;font-size:14px;'>Meeting Room Booking &mdash; " + CONFIG.ORG_NAME + "</p>" +
    "</div>" +

    "<div style='background:white;padding:28px 32px;border-radius:0 0 12px 12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);'>" +

      "<p style='color:#374151;font-size:15px;margin:0 0 20px;'>" +
        "Halo <strong>" + (data.namaLengkap || "Karyawan") + "</strong>,<br>" +
        "Maaf, booking ruang meeting Anda <strong style='color:#dc2626;'>tidak dapat diproses</strong> karena ruangan sudah dibooking pada waktu yang sama." +
      "</p>" +

      "<div style='background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:20px;margin-bottom:16px;'>" +
        "<h3 style='margin:0 0 12px;color:#991b1b;font-size:15px;'>Booking Anda (Ditolak)</h3>" +
        "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
          "<tr><td style='padding:6px 0;color:#6b7280;width:130px;'>Meeting</td><td style='padding:6px 0;color:#111827;font-weight:600;'>" + (data.judulMeeting||"-") + "</td></tr>" +
          "<tr><td style='padding:6px 0;color:#6b7280;'>Ruangan</td><td style='padding:6px 0;color:#111827;'>" + (data.namaRoom||"-") + "</td></tr>" +
          "<tr><td style='padding:6px 0;color:#6b7280;'>Tanggal</td><td style='padding:6px 0;color:#111827;'>" + formatTanggalIndo(data.tanggal) + "</td></tr>" +
          "<tr><td style='padding:6px 0;color:#6b7280;'>Waktu</td><td style='padding:6px 0;color:#dc2626;font-weight:700;'>" + (data.jamMulai||"-") + " &ndash; " + (data.jamSelesai||"-") + "</td></tr>" +
        "</table>" +
      "</div>" +

      "<div style='background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;padding:20px;margin-bottom:20px;'>" +
        "<h3 style='margin:0 0 12px;color:#9a3412;font-size:15px;'>Sudah Dibooking Oleh</h3>" +
        "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
          "<tr><td style='padding:6px 0;color:#6b7280;width:130px;'>Meeting</td><td style='padding:6px 0;color:#111827;font-weight:600;'>" + (conflict.judulMeeting||"-") + "</td></tr>" +
          "<tr><td style='padding:6px 0;color:#6b7280;'>Waktu</td><td style='padding:6px 0;color:#9a3412;font-weight:700;'>" + (conflict.jamMulai||"-") + " &ndash; " + (conflict.jamSelesai||"-") + "</td></tr>" +
        "</table>" +
      "</div>" +

      "<div style='background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:14px;margin-bottom:20px;'>" +
        "<p style='margin:0;font-size:13px;color:#1e40af;'>" +
          "<strong>Saran:</strong> Silakan pilih ruangan lain atau waktu yang berbeda, lalu isi kembali Google Form booking." +
        "</p>" +
      "</div>" +

      "<p style='font-size:13px;color:#6b7280;margin:0;'>" +
        "Pertanyaan? Hubungi <a href='mailto:" + CONFIG.EMAIL_ADMIN + "' style='color:#4f46e5;'>" + CONFIG.EMAIL_ADMIN + "</a>" +
      "</p>" +

    "</div></div>";

  GmailApp.sendEmail(data.email, subject, stripHtml(html), { htmlBody: html, name: "Meeting Room Booking " + CONFIG.ORG_NAME });
  Logger.log("Email gagal terkirim ke: " + data.email);
}

// ═══════════════════════════════════════════════════════════════════════════════
// NOTIFIKASI ADMIN — COLLISION TERDETEKSI
// ═══════════════════════════════════════════════════════════════════════════════
function sendEmailAdminCollision(data, conflict) {
  var subject = "[Booking Ditolak - Collision] " + (data.judulMeeting||"") + " - Meeting Room Booking " + CONFIG.ORG_NAME;

  var html =
    "<div style='font-family:Arial,sans-serif;max-width:600px;margin:0 auto;'>" +
      "<div style='background:#dc2626;padding:20px 24px;border-radius:10px 10px 0 0;'>" +
        "<h2 style='color:white;margin:0;font-size:18px;'>Booking Ditolak - Bentrok Jadwal</h2>" +
        "<p style='color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;'>Meeting Room Booking &mdash; " + CONFIG.ORG_NAME + "</p>" +
      "</div>" +
      "<div style='background:white;padding:24px;border:1px solid #e5e7eb;border-radius:0 0 10px 10px;'>" +
        "<p style='color:#374151;margin:0 0 16px;font-size:14px;'>Booking berikut ditolak karena ruangan sudah dibooking pada jam yang sama:</p>" +
        "<table style='width:100%;border-collapse:collapse;font-size:14px;'>" +
          "<tr style='background:#fef2f2;'><td style='padding:9px 12px;font-weight:600;color:#374151;width:150px;'>Pengaju</td><td style='padding:9px 12px;'>" + (data.namaLengkap||"-") + " (" + (data.email||"-") + ")</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Meeting Ditolak</td><td style='padding:9px 12px;'>" + (data.judulMeeting||"-") + "</td></tr>" +
          "<tr style='background:#fef2f2;'><td style='padding:9px 12px;font-weight:600;color:#374151;'>Ruangan</td><td style='padding:9px 12px;'>" + (data.namaRoom||"-") + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Tanggal</td><td style='padding:9px 12px;'>" + formatTanggalIndo(data.tanggal) + "</td></tr>" +
          "<tr style='background:#fef2f2;'><td style='padding:9px 12px;font-weight:600;color:#dc2626;'>Waktu Diminta</td><td style='padding:9px 12px;font-weight:700;color:#dc2626;'>" + (data.jamMulai||"-") + " &ndash; " + (data.jamSelesai||"-") + "</td></tr>" +
          "<tr><td style='padding:9px 12px;font-weight:600;color:#374151;'>Konflik Dengan</td><td style='padding:9px 12px;font-weight:700;color:#9a3412;'>" + (conflict.judulMeeting||"-") + " (" + (conflict.jamMulai||"-") + " &ndash; " + (conflict.jamSelesai||"-") + ")</td></tr>" +
        "</table>" +
      "</div>" +
    "</div>";

  GmailApp.sendEmail(CONFIG.EMAIL_ADMIN, subject, stripHtml(html), { htmlBody: html, name: "Meeting Room Booking System" });
  Logger.log("Notifikasi collision terkirim ke admin.");
}

// ═══════════════════════════════════════════════════════════════════════════════
// SIMPAN KE GOOGLE SHEET (bridge ke HTML app)
// Kolom harus PERSIS urutan ini agar HTML app bisa baca
// ═══════════════════════════════════════════════════════════════════════════════
function saveToSheet(data) {
  try {
    var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var tab = ss.getSheetByName(CONFIG.SHEET_TAB);

    // Buat tab baru jika belum ada
    if (!tab) {
      tab = ss.insertSheet(CONFIG.SHEET_TAB);
    }

    // Jika sheet masih kosong, buat header
    if (tab.getLastRow() === 0) {
      var header = [
        "id","timestamp","namaLengkap","email","departemen",
        "judulMeeting","namaRoom","roomId","tanggal",
        "jamMulai","jamSelesai","namaPeserta","jumlahPeserta",
        "fasilitas","catatan"
      ];
      tab.appendRow(header);
      tab.getRange(1, 1, 1, header.length)
         .setFontWeight("bold")
         .setBackground("#4f46e5")
         .setFontColor("white");
      tab.setFrozenRows(1);
    }

    var id = "BK-" + new Date().getTime();
    var newRow = tab.getLastRow() + 1;
    var rowData = [
      id,
      data.timestamp       || "",
      data.namaLengkap     || "",
      data.email           || "",
      data.departemen      || "",
      data.judulMeeting    || "",
      data.namaRoom        || "",
      data.roomId          || "",
      data.tanggal         || "",
      data.jamMulai        || "",
      data.jamSelesai      || "",
      data.namaPeserta     || "",
      data.jumlahPeserta   || "",
      data.fasilitas       || "",
      data.catatan         || "",
    ];
    // Set kolom jam (index 9 & 10, kolom J & K) sebagai Plain Text dulu
    // agar Sheets tidak auto-convert "09:00" → Date object
    tab.getRange(newRow, 10, 1, 2).setNumberFormat("@STRING@");
    tab.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    Logger.log("Data tersimpan di Sheet tab '" + CONFIG.SHEET_TAB + "', ID: " + id);
  } catch (err) {
    Logger.log("Gagal simpan ke Sheet: " + err.toString());
    // Kirim error ke admin agar tahu
    try {
      GmailApp.sendEmail(CONFIG.EMAIL_ADMIN, "[RoomBook] Gagal simpan ke Sheet", err.toString());
    } catch(e2) {}
  }
}

// ─── Helper: strip HTML tags untuk fallback plain text ───────────────────────
function stripHtml(html) {
  if (!html) return "";
  return String(html).replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
}

// ═══════════════════════════════════════════════════════════════════════════════
// TEST COLLISION — jalankan dari editor untuk debug apakah collision terdeteksi
// Sesuaikan tanggal/jam/room dengan booking yang sudah ada di sheet
// ═══════════════════════════════════════════════════════════════════════════════
function testCollision() {
  var fakeData = {
    email:         "zulfikar.dwi@majoo.id",
    namaLengkap:   "Test User",
    departemen:    "Test",
    judulMeeting:  "Test Collision",
    namaRoom:      "Raja Ampat",    // sama dengan booking yang ada di sheet
    tanggal:       "2026-05-06",    // sama dengan tanggal di sheet
    jamMulai:      "10:00",         // overlap dengan 9:00-11:00
    jamSelesai:    "11:00",
    jumlahPeserta: "2",
    namaPeserta:   "Test",
    fasilitas:     "",
    catatan:       "",
    timestamp:     new Date().toISOString(),
    roomId:        1,               // Raja Ampat = 1
    kapasitas:     7,
    fasilitasRoom: "TV, Whiteboard",
  };

  Logger.log("=== TEST COLLISION ===");
  Logger.log("Data: " + JSON.stringify(fakeData));

  var conflict = cekCollision(fakeData);
  if (conflict) {
    Logger.log("COLLISION TERDETEKSI: " + JSON.stringify(conflict));
    Logger.log("Kirim email gagal ke: " + fakeData.email);
    sendEmailGagal(fakeData, conflict);
    sendEmailAdminCollision(fakeData, conflict);
    Logger.log("Test selesai — cek inbox email");
  } else {
    Logger.log("Tidak ada collision ditemukan. Pastikan ada booking di RoomBook Data untuk room + tanggal + jam yang sama.");
    // Tampilkan isi sheet untuk debug
    try {
      var ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      var tab = ss.getSheetByName(CONFIG.SHEET_TAB);
      if (!tab) { Logger.log("Tab '" + CONFIG.SHEET_TAB + "' belum ada!"); return; }
      var values = tab.getDataRange().getValues();
      Logger.log("Isi sheet (" + values.length + " rows):");
      values.forEach(function(r, i) { Logger.log("Row " + i + ": " + JSON.stringify(r)); });
    } catch(e) { Logger.log("Gagal baca sheet: " + e); }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// TEST MANUAL — jalankan dari editor untuk mencoba tanpa isi form
// ═══════════════════════════════════════════════════════════════════════════════
function testManual() {
  var fakeData = {
    email:         "zulfikar.dwi@majoo.id",
    namaLengkap:   "Zulfikar Dwi",
    departemen:    "Product",
    judulMeeting:  "Test Booking dari Apps Script",
    namaRoom:      "Bromo",
    tanggal:       "2026-05-10",
    jamMulai:      "10:00",
    jamSelesai:    "11:30",
    jumlahPeserta: "3 – 4 orang",
    namaPeserta:   "Zulfikar, Firjayanti",
    fasilitas:     "TV, Camera Conference",
    catatan:       "Test manual dari editor",
    timestamp:     new Date().toISOString(),
    roomId:        2,
    kapasitas:     9,
    fasilitasRoom: "TV, Whiteboard, Camera Conference",
  };

  Logger.log("=== TEST MANUAL ===");
  Logger.log("Data: " + JSON.stringify(fakeData));

  sendEmailKonfirmasi(fakeData);
  sendEmailAdmin(fakeData);

  Logger.log("=== SELESAI — cek inbox email ===");
}
