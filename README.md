# ✊ Fistbook

> A simple, clean sign-in system with an Excel-based database for sign-in attempt logging.

![White & Blue UI](https://img.shields.io/badge/theme-white%20%26%20blue-1877f2?style=flat-square)
![Node.js](https://img.shields.io/badge/Node.js-≥16-339933?style=flat-square&logo=node.js)
![Excel DB](https://img.shields.io/badge/database-Excel%20.xlsx-217346?style=flat-square&logo=microsoftexcel)

---

## 📦 Features

- **Sign In** – Validates credentials against an Excel user database
- **Register** – Creates a new account (stored in `data/users.xlsx`)
- **Excel Logging** – Every sign-in attempt (success or failure) is logged to `data/signin_attempts.xlsx`
- **Password Security** – Passwords are hashed with bcrypt (never stored as plain text)
- **Password Strength Meter** – Live feedback on registration page
- **Responsive UI** – Works on desktop and mobile
- **Session Management** – Server-side sessions with express-session

---

## 🗂️ Project Structure

```
fistbook/
├── public/                 # Static frontend files
│   ├── index.html          # Sign-in page
│   ├── register.html       # Registration page
│   ├── dashboard.html      # Post-login dashboard
│   ├── style.css           # All styles (white + blue theme)
│   ├── auth.js             # Sign-in logic
│   ├── register.js         # Registration logic
│   └── dashboard.js        # Dashboard logic
├── data/                   # Auto-created on first run
│   ├── users.xlsx          # User accounts database
│   └── signin_attempts.xlsx# Sign-in attempt log
├── scripts/
│   ├── init_excel.py       # Pre-format Excel files (optional)
│   └── view_attempts.py    # CLI viewer for attempt logs
├── server.js               # Express backend (API + static server)
├── package.json
└── README.md
```

---

## 🚀 Quick Start

### 1. Install dependencies

```bash
npm install
```

### 2. (Optional) Pre-create formatted Excel files

```bash
pip install openpyxl
python scripts/init_excel.py
```

### 3. Start the server

```bash
npm start
```

Or for development with auto-reload:

```bash
npm run dev
```

### 4. Open in browser

```
http://localhost:3000
```

---

## 📊 Excel Database Files

Both files are auto-created in the `data/` folder when the server first runs.

### `data/users.xlsx` — User Accounts

| Column         | Description                       |
|----------------|-----------------------------------|
| ID             | Auto-increment user ID            |
| First Name     | User's first name                 |
| Last Name      | User's last name                  |
| Email          | Unique email address              |
| Username       | Unique username (lowercase)       |
| Password Hash  | bcrypt hash (never plain text)    |
| Created At     | Registration timestamp            |

### `data/signin_attempts.xlsx` — Attempt Log

| Column       | Description                              |
|--------------|------------------------------------------|
| #            | Attempt number                           |
| Timestamp    | Date and time of the attempt             |
| Username     | Username or email entered                |
| IP Address   | Client IP address                        |
| Result       | `SUCCESS` or `FAIL`                      |
| Reason       | e.g. "Credentials valid", "Wrong password"|
| User Agent   | Browser/client info                      |

---

## 🔌 API Endpoints

| Method | Endpoint           | Description               |
|--------|--------------------|---------------------------|
| POST   | `/api/register`    | Create a new account      |
| POST   | `/api/signin`      | Sign in                   |
| POST   | `/api/logout`      | Destroy session           |
| GET    | `/api/attempts`    | List all sign-in attempts |

### Register – Request Body
```json
{
  "firstName": "John",
  "lastName": "Doe",
  "email": "john@example.com",
  "username": "johndoe",
  "password": "MyPassword123!"
}
```

### Sign In – Request Body
```json
{
  "username": "johndoe",
  "password": "MyPassword123!"
}
```

---

## 🖥️ View Sign-In Attempts (CLI)

```bash
# Show last 50 attempts (default)
python scripts/view_attempts.py

# Show last 20 attempts
python scripts/view_attempts.py --last 20

# Show only failed attempts
python scripts/view_attempts.py --failed
```

---

## ⚙️ Environment Variables

| Variable          | Default                              | Description         |
|-------------------|--------------------------------------|---------------------|
| `PORT`            | `3000`                               | Server port         |
| `SESSION_SECRET`  | `fistbook-secret-change-in-production` | Session secret key  |

Set them in a `.env` file (install `dotenv` if needed) or export before running:

```bash
SESSION_SECRET=my-super-secret PORT=8080 npm start
```

---

## 🔒 Security Notes

- Passwords are **never stored in plain text** — bcrypt with 12 salt rounds is used.
- Change `SESSION_SECRET` to a random string in production.
- For production, set `cookie.secure = true` in `server.js` and serve over HTTPS.
- The `/api/attempts` endpoint has no authentication — add middleware to restrict access in production.

---

## 📄 License

MIT
