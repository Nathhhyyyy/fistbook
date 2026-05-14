// server.js – Fistbook Backend (Node.js + Express)
// Stores users and sign-in attempts in Excel (.xlsx)

const express    = require('express');
const path       = require('path');
const bcrypt     = require('bcryptjs');
const session    = require('express-session');
const ExcelJS    = require('exceljs');
const fs         = require('fs');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Paths ────────────────────────────────────────────────────────────────────
const DATA_DIR          = path.join(__dirname, 'data');
const USERS_FILE        = path.join(DATA_DIR, 'users.xlsx');
const ATTEMPTS_FILE     = path.join(DATA_DIR, 'signin_attempts.xlsx');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
  secret: process.env.SESSION_SECRET || 'fistbook-secret-change-in-production',
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 1000 * 60 * 60 * 24 }, // 24 h
}));

// ── Excel Helpers ─────────────────────────────────────────────────────────────

/** Ensure the users workbook exists with headers */
async function ensureUsersFile() {
  if (fs.existsSync(USERS_FILE)) return;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Users');
  ws.columns = [
    { header: 'ID',           key: 'id',         width: 10 },
    { header: 'First Name',   key: 'firstName',   width: 18 },
    { header: 'Last Name',    key: 'lastName',    width: 18 },
    { header: 'Email',        key: 'email',       width: 30 },
    { header: 'Username',     key: 'username',    width: 20 },
    { header: 'Password Hash',key: 'hash',        width: 70 },
    { header: 'Created At',   key: 'createdAt',   width: 22 },
  ];
  styleHeader(ws);
  await wb.xlsx.writeFile(USERS_FILE);
}

/** Ensure the attempts workbook exists with headers */
async function ensureAttemptsFile() {
  if (fs.existsSync(ATTEMPTS_FILE)) return;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sign-In Attempts');
  ws.columns = [
    { header: '#',            key: 'id',        width: 8  },
    { header: 'Timestamp',    key: 'timestamp', width: 24 },
    { header: 'Username',     key: 'username',  width: 22 },
    { header: 'IP Address',   key: 'ip',        width: 18 },
    { header: 'Result',       key: 'result',    width: 14 },
    { header: 'Reason',       key: 'reason',    width: 35 },
    { header: 'User Agent',   key: 'userAgent', width: 60 },
  ];
  styleHeader(ws);
  await wb.xlsx.writeFile(ATTEMPTS_FILE);
}

/** Apply header row styling */
function styleHeader(ws) {
  const headerRow = ws.getRow(1);
  headerRow.eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1877F2' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FF1563CC' } },
    };
  });
  headerRow.height = 22;
}

/** Load all rows from a workbook sheet as objects */
async function loadRows(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  const headers = [];
  ws.getRow(1).eachCell(cell => headers.push(cell.value));
  const rows = [];
  ws.eachRow((row, idx) => {
    if (idx === 1) return;
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colIdx) => {
      obj[headers[colIdx - 1]] = cell.value;
    });
    rows.push(obj);
  });
  return { wb, ws, rows, headers };
}

/** Append a sign-in attempt row to the attempts Excel file */
async function logAttempt({ username, ip, result, reason, userAgent }) {
  await ensureAttemptsFile();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(ATTEMPTS_FILE);
  const ws = wb.worksheets[0];

  const rowCount = ws.rowCount;   // includes header
  const newId    = rowCount;      // row 1 = header, so id = rowCount for next data row

  const newRow = ws.addRow({
    id:        newId,
    timestamp: new Date().toLocaleString('en-US', { hour12: true }),
    username,
    ip,
    result,
    reason,
    userAgent: (userAgent || '').substring(0, 120),
  });

  // Alternate row shading
  const isEven = (newId % 2 === 0);
  newRow.eachCell({ includeEmpty: true }, cell => {
    cell.fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: isEven ? 'FFF0F7FF' : 'FFFFFFFF' },
    };
    cell.font = { name: 'Arial', size: 10 };
    cell.alignment = { vertical: 'middle' };
    // Color-code result column
    if (cell.col === 5) {
      cell.font = {
        name: 'Arial', size: 10, bold: true,
        color: { argb: result === 'SUCCESS' ? 'FF38A169' : 'FFE53E3E' },
      };
    }
  });
  newRow.height = 18;

  await wb.xlsx.writeFile(ATTEMPTS_FILE);
}

// ── API Routes ────────────────────────────────────────────────────────────────

/** POST /api/register */
app.post('/api/register', async (req, res) => {
  try {
    const { firstName, lastName, email, username, password } = req.body;

    // Basic validation
    if (!firstName || !lastName || !email || !username || !password) {
      return res.status(400).json({ success: false, message: 'All fields are required.' });
    }
    if (password.length < 8) {
      return res.status(400).json({ success: false, message: 'Password must be at least 8 characters.' });
    }

    await ensureUsersFile();
    const { wb, ws, rows } = await loadRows(USERS_FILE);

    // Check duplicates
    const emailLower    = email.toLowerCase().trim();
    const usernameLower = username.toLowerCase().trim();

    const duplicate = rows.find(r =>
      String(r['Email']   || '').toLowerCase() === emailLower ||
      String(r['Username']|| '').toLowerCase() === usernameLower
    );
    if (duplicate) {
      const field = String(duplicate['Email'] || '').toLowerCase() === emailLower ? 'email' : 'username';
      return res.status(409).json({ success: false, message: `This ${field} is already registered.` });
    }

    const hash = await bcrypt.hash(password, 12);
    const newId = rows.length + 1;

    const newRow = ws.addRow({
      id:        newId,
      firstName: firstName.trim(),
      lastName:  lastName.trim(),
      email:     emailLower,
      username:  usernameLower,
      hash,
      createdAt: new Date().toLocaleString('en-US', { hour12: true }),
    });

    const isEven = (newId % 2 === 0);
    newRow.eachCell({ includeEmpty: true }, cell => {
      cell.fill = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: isEven ? 'FFF0F7FF' : 'FFFFFFFF' },
      };
      cell.font = { name: 'Arial', size: 10 };
      cell.alignment = { vertical: 'middle' };
    });
    newRow.height = 18;

    await wb.xlsx.writeFile(USERS_FILE);
    return res.json({ success: true, message: 'Account created successfully.' });

  } catch (err) {
    console.error('Register error:', err);
    return res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
});

/** POST /api/signin */
app.post('/api/signin', async (req, res) => {
  const ip        = req.headers['x-forwarded-for'] || req.socket.remoteAddress || 'unknown';
  const userAgent = req.headers['user-agent'] || '';
  const { username, password } = req.body;

  try {
    if (!username || !password) {
      await logAttempt({ username: username || '', ip, result: 'FAIL', reason: 'Missing credentials', userAgent });
      return res.status(400).json({ success: false, message: 'Username and password are required.' });
    }

    await ensureUsersFile();
    const { rows } = await loadRows(USERS_FILE);

    const user = rows.find(r =>
      String(r['Username'] || '').toLowerCase() === username.toLowerCase().trim() ||
      String(r['Email']    || '').toLowerCase() === username.toLowerCase().trim()
    );

    if (!user) {
      await logAttempt({ username, ip, result: 'FAIL', reason: 'User not found', userAgent });
      return res.status(401).json({ success: false, message: 'Invalid username or password.' });
    }

    const match = await bcrypt.compare(password, String(user['Password Hash']));

    if (!match) {
      await logAttempt({ username, ip, result: 'FAIL', reason: 'Wrong password', userAgent });
      return res.status(401).json({ success: false, message: 'Invalid username or password.' });
    }

    // Success – set session
    req.session.userId   = user['ID'];
    req.session.username = user['Username'];

    await logAttempt({ username, ip, result: 'SUCCESS', reason: 'Credentials valid', userAgent });

    return res.json({
      success:   true,
      message:   'Signed in successfully.',
      username:  user['Username'],
      firstName: user['First Name'],
    });

  } catch (err) {
    console.error('Signin error:', err);
    await logAttempt({ username: username || '', ip, result: 'FAIL', reason: `Server error: ${err.message}`, userAgent });
    return res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
});

/** POST /api/logout */
app.post('/api/logout', (req, res) => {
  req.session.destroy(() => {
    res.json({ success: true });
  });
});

/** GET /api/attempts – View sign-in attempts (for admins) */
app.get('/api/attempts', async (req, res) => {
  try {
    await ensureAttemptsFile();
    const { rows } = await loadRows(ATTEMPTS_FILE);
    return res.json({ success: true, attempts: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Could not load attempts.' });
  }
});

// ── Start Server ──────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✊ Fistbook running at http://localhost:${PORT}\n`);
  // Pre-create the data files on startup
  ensureUsersFile().catch(console.error);
  ensureAttemptsFile().catch(console.error);
});
