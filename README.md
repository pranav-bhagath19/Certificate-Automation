# 🎓 Certificate Automation

Automatically generate personalized certificates and email them to 500+ participants in minutes. Built for **STRAT-A-THON 1.0** at Vishnu Institute of Technology, organized by the Techie Blazers Club, CSBS Department.

---

## ✨ Features

- 📊 Reads participant **Name** and **Email** directly from Excel
- 🎨 Overlays each participant's name onto your **Canva-designed certificate**
- 📄 Converts every certificate to a **PDF automatically**
- 📧 Sends personalized emails with certificate attached via **Gmail**
- ⚡ **Parallel processing** — sends 10 certificates simultaneously
- 💾 **Auto-saves progress** — safely resume if interrupted
- ⏭️ **Skips already-sent** — no duplicate emails on re-run
- 📁 Saves all PDFs locally in `certificates_output/`

---

## 📁 Project Structure

```
certificate-automation/
│
├── certificates_fast.py      # Main script (fast parallel version)
├── certificates.py           # Simple single-threaded version
├── certificate_bg.png        # Your certificate design (not committed)
├── participants.xlsx         # Excel file with Name & Email (not committed)
├── progress.json             # Auto-generated — tracks sent emails
├── certificates_output/      # Auto-generated — all PDFs saved here
├── .gitignore
└── README.md
```

---

## 🚀 Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/yourusername/certificate-automation.git
cd certificate-automation
```

### 2. Install dependencies
```bash
pip install pillow openpyxl reportlab tqdm
```

### 3. Add your files
Place these two files in the project folder:
- `certificate_bg.png` — your Canva certificate exported as PNG (without the name)
- `participants.xlsx` — Excel file with columns `Name` and `Email`

### 4. Set up Gmail App Password
1. Go to [myaccount.google.com](https://myaccount.google.com)
2. **Security → 2-Step Verification** → enable it
3. **Security → App Passwords** → select Mail → Generate
4. Copy the 16-character password

### 5. Configure the script
Open `certificates_fast.py` and edit the top section:
```python
GMAIL_ADDRESS  = "youremail@gmail.com"
GMAIL_APP_PASS = "xxxx xxxx xxxx xxxx"   # your App Password
```

### 6. Run
```bash
python certificates_fast.py
```

---

## 📊 Excel Format

Your `participants.xlsx` must have these exact column headers in Row 1:

| Name | Email |
|------|-------|
| Ravi Kumar | ravi@gmail.com |
| Priya Sharma | priya@gmail.com |

---

## ⚙️ Configuration

All settings are at the top of `certificates_fast.py`:

| Setting | Default | Description |
|---------|---------|-------------|
| `EXCEL_FILE` | `participants.xlsx` | Path to your Excel file |
| `CERTIFICATE_BG` | `certificate_bg.png` | Path to your certificate PNG |
| `OUTPUT_FOLDER` | `certificates_output` | Where PDFs are saved |
| `WORKERS` | `10` | Parallel certificate generators |
| `SMTP_CONNECTIONS` | `5` | Parallel Gmail senders |
| `BATCH_SIZE` | `50` | Save progress every N certificates |
| `NAME_VERTICAL_POSITION` | `0.396` | Name Y position (0.0–1.0) |
| `NAME_HORIZONTAL_POSITION` | `0.50` | Name X position (0.0–1.0) |
| `FONT_SIZE` | `80` | Name font size in pixels |
| `FONT_COLOR` | `(255,255,255)` | Name font color (RGB) |

---

## ⚡ Speed

| Participants | Estimated Time |
|-------------|----------------|
| 100 | ~30 seconds |
| 500 | ~2–3 minutes |
| 1000 | ~5–6 minutes |

> Speed depends on your internet connection and Gmail rate limits.

---

## 🔁 Resume After Interruption

If the script stops mid-way (power cut, internet drop), just re-run it:

```bash
python certificates_fast.py
```

It reads `progress.json` and **automatically skips** everyone who already received their certificate.

To start completely fresh, delete `progress.json`:
```bash
del progress.json       # Windows
rm progress.json        # Mac / Linux
```

---

## ⚠️ Gmail Daily Limit

| Account Type | Daily Limit |
|-------------|-------------|
| Free Gmail | 500 emails/day |
| Google Workspace | 2,000 emails/day |

For 500 participants, a free Gmail account is sufficient.

---

## 🛡️ Security

- **Never commit** your `participants.xlsx` — it contains personal data
- **Never commit** your Gmail App Password
- Both are already covered in `.gitignore`
- Consider using environment variables for credentials:

```python
import os
GMAIL_APP_PASS = os.environ.get("GMAIL_APP_PASS")
```

Then set it in your terminal:
```bash
# Windows
set GMAIL_APP_PASS=xxxx xxxx xxxx xxxx

# Mac / Linux
export GMAIL_APP_PASS="xxxx xxxx xxxx xxxx"
```

---

## 🛠️ Troubleshooting

| Problem | Fix |
|---------|-----|
| `Gmail authentication failed` | Use App Password, not your regular Gmail password |
| `Name column not found` | Make sure Row 1 headers are exactly `Name` and `Email` |
| Name in wrong position | Adjust `NAME_VERTICAL_POSITION` (try `±0.02` increments) |
| Font looks wrong | Script auto-detects best available system font |
| PDFs not generating | Run `pip install reportlab` |
| Script crashes mid-way | Just re-run — progress is saved automatically |

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `Pillow` | Image processing — overlays name on certificate |
| `openpyxl` | Reads Excel files |
| `reportlab` | Converts image to PDF |
| `tqdm` | Live progress bar |

Install all at once:
```bash
pip install pillow openpyxl reportlab tqdm
# 📄 License

MIT License — free to use and modify for your own events.
