/**
 * RoomBook — Backend API (Node.js / Express)
 * ============================================
 * Endpoints:
 *   POST   /api/auth/login
 *   GET    /api/rooms
 *   GET    /api/bookings
 *   POST   /api/bookings
 *   PUT    /api/bookings/:id
 *   DELETE /api/bookings/:id
 *   POST   /api/bookings/form-submit   ← dari Google Form via Apps Script
 *
 * Run:
 *   npm install
 *   node server.js
 *
 * Environment variables (.env):
 *   PORT=3001
 *   JWT_SECRET=your-secret-key
 *   EMAIL_HOST=smtp.gmail.com
 *   EMAIL_PORT=587
 *   EMAIL_USER=your@email.com
 *   EMAIL_PASS=your-app-password
 *   FORM_API_KEY=roombook-form-key-2024   ← API key untuk Google Form
 */

require("dotenv").config();
const express    = require("express");
const cors       = require("cors");
const jwt        = require("jsonwebtoken");
const bcrypt     = require("bcryptjs");
const nodemailer = require("nodemailer");

const app    = express();
const PORT   = process.env.PORT || 3001;
const SECRET = process.env.JWT_SECRET || "roombook-secret-dev";

app.use(cors());
app.use(express.json());

// ─── IN-MEMORY DATABASE (replace with real DB in production) ──────────────────
const db = {
  users: [
    { id: 1, name: "Zulfikar Dwi", email: "zulfikar.dwi@majoo.id", passwordHash: bcrypt.hashSync("admin123", 10), role: "Admin",    avatar: "ZD" },
    { id: 2, name: "Firjayanti",   email: "firjayanti@majoo.id",    passwordHash: bcrypt.hashSync("pass123",  10), role: "Employee", avatar: "FJ" },
  ],
  rooms: [
    { id: 1, name: "Raja Ampat",     capacity: 7, colorHex: "#3b82f6", facilities: ["TV", "Whiteboard"] },
    { id: 2, name: "Bromo",          capacity: 9, colorHex: "#10b981", facilities: ["TV", "Whiteboard", "Camera Conference"] },
    { id: 3, name: "Tegallalang",    capacity: 6, colorHex: "#8b5cf6", facilities: ["TV", "Whiteboard", "Camera Conference"] },
    { id: 4, name: "Tanjung Tinggi", capacity: 6, colorHex: "#f59e0b", facilities: ["TV", "Whiteboard"] },
  ],
  bookings: [
    { id: 1, roomId: 1, title: "Q2 Planning",     date: offsetDate(0), startTime: "09:00", endTime: "10:30", attendees: "Product Team",  notes: "", bookedById: 1, bookedByName: "Zulfikar Dwi", createdAt: new Date().toISOString() },
    { id: 2, roomId: 2, title: "1-on-1 Review",   date: offsetDate(0), startTime: "11:00", endTime: "12:00", attendees: "HR + Zulfikar", notes: "", bookedById: 2, bookedByName: "Firjayanti",    createdAt: new Date().toISOString() },
  ],
  nextId: 100,
};

function offsetDate(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

// ─── MIDDLEWARE: JWT AUTH ─────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  const auth = req.headers.authorization;
  if (!auth || !auth.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Token tidak ditemukan." });
  }
  try {
    req.user = jwt.verify(auth.slice(7), SECRET);
    next();
  } catch {
    res.status(401).json({ error: "Token tidak valid atau kadaluarsa." });
  }
}

function requireAdmin(req, res, next) {
  if (req.user?.role !== "Admin") {
    return res.status(403).json({ error: "Akses ditolak. Hanya Admin." });
  }
  next();
}

// ─── NOTIFICATIONS ─────────────────────────────────────────────────────────────
const emailTransporter = nodemailer.createTransport({
  host:   process.env.EMAIL_HOST || "smtp.gmail.com",
  port:   parseInt(process.env.EMAIL_PORT) || 587,
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

async function sendEmailNotification({ to, subject, html }) {
  if (!process.env.EMAIL_USER) {
    console.log(`[EMAIL SIMULATED] To: ${to} | Subject: ${subject}`);
    return;
  }
  try {
    await emailTransporter.sendMail({ from: `"RoomBook" <${process.env.EMAIL_USER}>`, to, subject, html });
    console.log(`[EMAIL SENT] To: ${to}`);
  } catch (err) {
    console.error("[EMAIL ERROR]", err.message);
  }
}

async function notifyBooking(type, booking, user) {
  const room  = db.rooms.find(r => r.id === booking.roomId);
  const emojis = { create: "✅", update: "✏️", delete: "🗑️" };
  const emailHtml = `
    <h2>${emojis[type]} Booking ${type === "create" ? "Baru" : type === "update" ? "Diperbarui" : "Dihapus"}</h2>
    <table style="border-collapse:collapse;width:100%">
      <tr><td style="padding:6px 12px;font-weight:bold">Meeting</td><td style="padding:6px 12px">${booking.title}</td></tr>
      <tr><td style="padding:6px 12px;font-weight:bold">Room</td><td style="padding:6px 12px">${room?.name}</td></tr>
      <tr><td style="padding:6px 12px;font-weight:bold">Tanggal</td><td style="padding:6px 12px">${booking.date}</td></tr>
      <tr><td style="padding:6px 12px;font-weight:bold">Waktu</td><td style="padding:6px 12px">${booking.startTime} – ${booking.endTime}</td></tr>
      <tr><td style="padding:6px 12px;font-weight:bold">Dibooking oleh</td><td style="padding:6px 12px">${user.name}</td></tr>
    </table>
  `;
  await sendEmailNotification({ to: user.email, subject: `[RoomBook] ${booking.title} — ${room?.name}`, html: emailHtml });
}

// ─── CONFLICT CHECK HELPER ────────────────────────────────────────────────────
function checkConflict(candidate, excludeId = null) {
  const cS = toMin(candidate.startTime), cE = toMin(candidate.endTime);
  return db.bookings.find(b => {
    if (b.id === excludeId) return false;
    if (b.roomId !== candidate.roomId || b.date !== candidate.date) return false;
    return cS < toMin(b.endTime) && cE > toMin(b.startTime);
  });
}
function toMin(t) { const [h, m] = t.split(":").map(Number); return h * 60 + m; }

// ─── ROUTES ───────────────────────────────────────────────────────────────────

// Health check
app.get("/api/health", (req, res) => res.json({ status: "ok", time: new Date().toISOString() }));

// POST /api/auth/login
app.post("/api/auth/login", (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) return res.status(400).json({ error: "Email dan password wajib diisi." });

  const user = db.users.find(u => u.email === email);
  if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
    return res.status(401).json({ error: "Email atau password salah." });
  }

  const token = jwt.sign(
    { id: user.id, email: user.email, name: user.name, role: user.role, avatar: user.avatar },
    SECRET, { expiresIn: "8h" }
  );
  res.json({ token, user: { id: user.id, name: user.name, email: user.email, role: user.role, avatar: user.avatar } });
});

// GET /api/rooms
app.get("/api/rooms", requireAuth, (req, res) => {
  res.json(db.rooms);
});

// GET /api/bookings?date=&roomId=
app.get("/api/bookings", requireAuth, (req, res) => {
  let result = [...db.bookings];
  if (req.query.date)   result = result.filter(b => b.date   === req.query.date);
  if (req.query.roomId) result = result.filter(b => b.roomId === parseInt(req.query.roomId));
  res.json(result);
});

// POST /api/bookings — create new booking
app.post("/api/bookings", requireAuth, async (req, res) => {
  const { roomId, title, date, startTime, endTime, attendees, notes } = req.body;

  // Validation
  if (!roomId || !title || !date || !startTime || !endTime) {
    return res.status(400).json({ error: "Field wajib: roomId, title, date, startTime, endTime." });
  }
  if (toMin(endTime) <= toMin(startTime)) {
    return res.status(400).json({ error: "Waktu selesai harus lebih dari waktu mulai." });
  }
  if (!db.rooms.find(r => r.id === roomId)) {
    return res.status(404).json({ error: "Room tidak ditemukan." });
  }

  // Conflict detection
  const conflict = checkConflict({ roomId, date, startTime, endTime });
  if (conflict) {
    return res.status(409).json({
      error: "Konflik jadwal.",
      conflict: { title: conflict.title, startTime: conflict.startTime, endTime: conflict.endTime, bookedByName: conflict.bookedByName },
    });
  }

  const booking = {
    id: db.nextId++, roomId, title, date, startTime, endTime,
    attendees: attendees || "", notes: notes || "",
    bookedById: req.user.id, bookedByName: req.user.name,
    createdAt: new Date().toISOString(),
  };
  db.bookings.push(booking);

  // Send notifications async (don't block response)
  notifyBooking("create", booking, req.user).catch(console.error);

  res.status(201).json(booking);
});

// PUT /api/bookings/:id — update booking
app.put("/api/bookings/:id", requireAuth, async (req, res) => {
  const id      = parseInt(req.params.id);
  const idx     = db.bookings.findIndex(b => b.id === id);
  if (idx === -1) return res.status(404).json({ error: "Booking tidak ditemukan." });

  const existing = db.bookings[idx];

  // Permission: owner or Admin
  if (existing.bookedById !== req.user.id && req.user.role !== "Admin") {
    return res.status(403).json({ error: "Hanya pembuat booking atau Admin yang bisa mengedit." });
  }

  const updated = { ...existing, ...req.body, id, bookedById: existing.bookedById, bookedByName: existing.bookedByName };

  if (toMin(updated.endTime) <= toMin(updated.startTime)) {
    return res.status(400).json({ error: "Waktu selesai harus lebih dari waktu mulai." });
  }

  const conflict = checkConflict(updated, id);
  if (conflict) {
    return res.status(409).json({
      error: "Konflik jadwal.",
      conflict: { title: conflict.title, startTime: conflict.startTime, endTime: conflict.endTime, bookedByName: conflict.bookedByName },
    });
  }

  db.bookings[idx] = updated;
  notifyBooking("update", updated, req.user).catch(console.error);
  res.json(updated);
});

// DELETE /api/bookings/:id
app.delete("/api/bookings/:id", requireAuth, async (req, res) => {
  const id  = parseInt(req.params.id);
  const idx = db.bookings.findIndex(b => b.id === id);
  if (idx === -1) return res.status(404).json({ error: "Booking tidak ditemukan." });

  const booking = db.bookings[idx];

  if (booking.bookedById !== req.user.id && req.user.role !== "Admin") {
    return res.status(403).json({ error: "Hanya pembuat booking atau Admin yang bisa menghapus." });
  }

  db.bookings.splice(idx, 1);
  notifyBooking("delete", booking, req.user).catch(console.error);
  res.json({ success: true, message: "Booking berhasil dihapus." });
});

// ─── POST /api/bookings/form-submit — dari Google Form via Apps Script ─────────
app.post("/api/bookings/form-submit", async (req, res) => {
  // Validasi API key (bukan JWT — Apps Script pakai header x-api-key)
  const apiKey = req.headers["x-api-key"];
  if (!apiKey || apiKey !== (process.env.FORM_API_KEY || "roombook-form-key-2024")) {
    return res.status(401).json({ error: "API key tidak valid." });
  }

  const { roomId, title, date, startTime, endTime, attendees, notes, bookedByName, bookedByEmail, source } = req.body;

  // Validasi field wajib
  if (!roomId || !title || !date || !startTime || !endTime) {
    return res.status(400).json({ error: "Field wajib: roomId, title, date, startTime, endTime." });
  }
  if (toMin(endTime) <= toMin(startTime)) {
    return res.status(400).json({ error: "Waktu selesai harus lebih dari waktu mulai." });
  }
  if (!db.rooms.find(r => r.id === Number(roomId))) {
    return res.status(404).json({ error: "Room tidak ditemukan." });
  }

  // Conflict detection
  const candidate = { roomId: Number(roomId), date, startTime, endTime };
  const conflict  = checkConflict(candidate);
  if (conflict) {
    return res.status(409).json({
      error: "Konflik jadwal.",
      conflict: {
        title:         conflict.title,
        startTime:     conflict.startTime,
        endTime:       conflict.endTime,
        bookedByName:  conflict.bookedByName,
      },
    });
  }

  // Simpan booking
  const booking = {
    id:            db.nextId++,
    roomId:        Number(roomId),
    title,
    date,
    startTime,
    endTime,
    attendees:     attendees     || "",
    notes:         notes         || "",
    bookedById:    null,          // dari form, bukan user login
    bookedByName:  bookedByName  || "Google Form",
    bookedByEmail: bookedByEmail || "",
    source:        source        || "google-form",
    createdAt:     new Date().toISOString(),
  };
  db.bookings.push(booking);

  // Kirim email notifikasi ke admin
  const room = db.rooms.find(r => r.id === booking.roomId);
  const emailHtml = `
    <h2>📅 Booking Baru via Google Form</h2>
    <table style="border-collapse:collapse;width:100%;font-family:sans-serif">
      <tr style="background:#f3e5f5"><td style="padding:8px 12px;font-weight:bold">Meeting</td><td style="padding:8px 12px">${booking.title}</td></tr>
      <tr><td style="padding:8px 12px;font-weight:bold">Ruangan</td><td style="padding:8px 12px">${room?.name}</td></tr>
      <tr style="background:#f3e5f5"><td style="padding:8px 12px;font-weight:bold">Tanggal</td><td style="padding:8px 12px">${booking.date}</td></tr>
      <tr><td style="padding:8px 12px;font-weight:bold">Waktu</td><td style="padding:8px 12px">${booking.startTime} – ${booking.endTime}</td></tr>
      <tr style="background:#f3e5f5"><td style="padding:8px 12px;font-weight:bold">Dibooking oleh</td><td style="padding:8px 12px">${booking.bookedByName} (${booking.bookedByEmail})</td></tr>
      <tr><td style="padding:8px 12px;font-weight:bold">Peserta</td><td style="padding:8px 12px">${booking.attendees || "-"}</td></tr>
      <tr style="background:#f3e5f5"><td style="padding:8px 12px;font-weight:bold">Catatan</td><td style="padding:8px 12px">${booking.notes || "-"}</td></tr>
    </table>
  `;

  const fakeUser = { email: process.env.EMAIL_USER || "admin@majoo.id", name: bookedByName || "Google Form" };
  sendEmailNotification({
    to:      fakeUser.email,
    subject: `[RoomBook] Booking Baru via Form — ${title}`,
    html:    emailHtml,
  }).catch(console.error);

  console.log(`[FORM SUBMIT] Booking #${booking.id} — ${title} oleh ${bookedByName}`);
  res.status(201).json(booking);
});

// ─── START SERVER ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n🚀 RoomBook API berjalan di http://localhost:${PORT}`);
  console.log(`   GET  /api/health`);
  console.log(`   POST /api/auth/login`);
  console.log(`   GET  /api/rooms`);
  console.log(`   GET  /api/bookings`);
  console.log(`   POST /api/bookings`);
  console.log(`   PUT  /api/bookings/:id`);
  console.log(`   DELETE /api/bookings/:id\n`);
});
