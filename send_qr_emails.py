"""
send_qr_emails.py  —  Powered by Brevo (free, safe, 600 students)
===================================================================

WHY BREVO INSTEAD OF GMAIL:
  Gmail flags bulk sends as spam and locks accounts.
  Brevo is built for bulk email — free tier: 300 emails/day.

  For 600 students:
    Day 1 → run script → sends 1 to 300 automatically
    Day 2 → run same script → sends 301 to 600 (resumes from log)
    No duplicates. No manual tracking needed.

SETUP (5 minutes, totally free):
  1. Go to https://app.brevo.com  and sign up free
  2. Go to Settings → SMTP & API → SMTP tab
  3. Copy your Login email and generate an SMTP Key
  4. Go to Settings → SMTP & API → API Keys tab → generate an API key
  5. Create a file called .env in the same folder as this script:

       BREVO_EMAIL=you@youremail.com
       BREVO_SMTP_KEY=xsmtpsib-xxxxxxxxxxxxxxxxxxxx
       BREVO_API_KEY=xkeysib-xxxxxxxxxxxxxxxxxxxx

  6. Install libraries:
       pip install pandas openpyxl qrcode pillow python-dotenv requests

  7. Run:
       python send_qr_emails.py

OUTPUT FILES:
  email_run.log          — full detailed log of every run
  sent_log.txt           — tracks sent emails for resume support
  failed_log.csv         — invalid/failed rows with reasons
  bounced_recipients.csv — emails that hard-bounced after sending
"""

import os, io, time, csv, logging, smtplib, traceback, requests
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.image     import MIMEImage
from datetime             import datetime, timedelta

import pandas as pd
import qrcode
from dotenv import load_dotenv
load_dotenv()

# ─────────────────────────────────────────────
#  CONFIG — edit these
# ─────────────────────────────────────────────
EXCEL_FILE    = "students.xlsx"
SHEET_NAME    = 0                   # 0 = first sheet

COL_NAME      = "Student Name"      # exact column headers in your Excel
COL_ROLL      = "Roll Number"
COL_EMAIL     = "Email"

SENDER_EMAIL  = os.getenv("BREVO_EMAIL", "hamzaahmed3956@outlook.com")
SENDER_NAME   = "HamzaX SecureEntry"
SENDER_PASS   = os.getenv("BREVO_SMTP_KEY", "")
BREVO_API_KEY = os.getenv("BREVO_API_KEY", "")  
SMTP_LOGIN   = os.getenv("SMTP_LOGIN", "")

SMTP_HOST     = "smtp-relay.brevo.com"
SMTP_PORT     = 587

SUBJECT       = "Your Welcome Fest 2025 — Official Entry Pass"
DAILY_LIMIT   = 300           # Brevo free = 300/day. Change to 500 on paid plan.
DELAY_SECONDS = 1.0           # 1 second between emails — safe and stable

# How many days back to check for bounces (7 is a good default)
BOUNCE_LOOKBACK_DAYS = 7

LOG_FILE      = "sent_log.txt"
FAILED_FILE   = "failed_log.csv"
BOUNCED_FILE  = "bounced_recipients.csv"
RUN_LOG_FILE  = "email_run.log"
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.FileHandler(RUN_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)


def load_sent_log() -> set:
    if not os.path.exists(LOG_FILE):
        return set()
    with open(LOG_FILE, encoding="utf-8") as f:
        return {line.strip().lower() for line in f if line.strip()}

def mark_sent(addr: str):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(addr.strip().lower() + "\n")

def generate_qr_bytes(data: str) -> bytes:
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8, border=3,
    )
    qr.add_data(str(data))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def build_email(recipient: str, name: str, roll: str, qr_bytes: bytes) -> MIMEMultipart:
    msg = MIMEMultipart("related")
    msg["From"]             = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"]               = recipient
    msg["Subject"]          = SUBJECT
    msg["List-Unsubscribe"] = f"<mailto:{SENDER_EMAIL}?subject=unsubscribe>"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Welcome Fest 2025 — Entry Pass</title>
</head>
<body style="margin:0;padding:0;background-color:#05050a;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="background-color:#05050a;min-height:100vh;">
  <tr>
    <td align="center" style="padding:40px 16px;">
      <table width="650" cellpadding="0" cellspacing="0" border="0"
             style="max-width:650px;width:100%;background-color:#0d0d16;
                    border:1px solid #1a1a2e;border-radius:16px;
                    overflow:hidden;box-shadow:0 0 60px rgba(0,255,136,0.08);">
        <tr>
          <td style="padding:0;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background:linear-gradient(135deg,#0a0a14 0%,#0f1628 50%,#0a0a14 100%);
                          border-bottom:1px solid #1a2a1a;">
              <tr>
                <td style="padding:36px 40px 28px 40px;text-align:center;">
                  <div style="height:2px;background:linear-gradient(90deg,transparent,#00ff88,#0088ff,transparent);margin-bottom:24px;"></div>
                  <p style="margin:0 0 6px 0;font-size:15px;letter-spacing:4px;
                             color:#00ff88;text-transform:uppercase;font-weight:700;">
                    UNIVERSITY OF ENGINEERING &amp; TECHNOLOGY, LAHORE
                  </p>
                  <p style="margin:0 0 18px 0;font-size:14px;letter-spacing:2px;
                             color:#7fffaa;text-transform:uppercase;">
                    SESSION 2024 PRESENTS
                  </p>
                  <h1 style="margin:0;font-size:52px;font-weight:900;letter-spacing:-1px;line-height:1.1;">
                    <span style="color:#ffffff;">WELCOME</span>
                    <span style="color:#00ff88;"> FEST</span>
                    <br/>
                    <span style="color:#00c8ff;font-size:58px;letter-spacing:2px;">2025</span>
                  </h1>
                  <p style="margin:14px 0 0 0;font-size:16px;letter-spacing:3px;
                             color:#c0caf5;text-transform:uppercase;">
                    &lt;YOUR OFFICIAL ENTRY PASS/&gt;
                  </p>
                  <div style="height:2px;background:linear-gradient(90deg,transparent,#0088ff,#00ff88,transparent);margin-top:24px;"></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:0;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#080810;border-bottom:1px solid #1a1a2e;">
              <tr>
                <td width="50%" style="padding:18px 24px;border-right:1px solid #1a1a2e;text-align:center;">
                  <p style="margin:0;font-size:14px;letter-spacing:3px;color:#00ff88;text-transform:uppercase;margin-bottom:4px;">
                    \U0001f4c5 DATE
                  </p>
                  <p style="margin:0;font-size:20px;font-weight:700;color:#ffffff;letter-spacing:1px;">
                    6 &amp; 7 MAY 2026
                  </p>
                </td>
                <td width="50%" style="padding:18px 24px;text-align:center;">
                  <p style="margin:0;font-size:14px;letter-spacing:3px;color:#00c8ff;text-transform:uppercase;margin-bottom:4px;">
                    \U0001f4cd VENUE
                  </p>
                  <p style="margin:0;font-size:18px;font-weight:700;color:#ffffff;letter-spacing:0.5px;">
                    Ground, Main Auditorium
                  </p>
                  <p style="margin:4px 0 0 0;font-size:14px;color:#a8b2d1;">UET Lahore Campus</p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:32px 40px 20px 40px;">
            <h2 style="margin:0;font-size:36px;font-weight:800;color:#ffffff;line-height:1.2;">
              Hello, {name}!
            </h2>
            <p style="margin:14px 0 0 0;font-size:18px;color:#ccd6f6;line-height:1.7;">
              Your <span style="color:#00ff88;font-weight:700;">personal entry pass</span>
              for Welcome Fest 2025 has been generated and is ready to use.
              Present this QR code at the entrance gate for instant verification.
            </p>
          </td>
        </tr>
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#0a0a18;border:1px solid #2a2a4a;
                          border-radius:10px;overflow:hidden;">
              <tr>
                <td>
                  <div style="background:linear-gradient(90deg,#00ff8820,#0088ff20);
                               padding:12px 20px;border-bottom:1px solid #2a2a4a;">
                    <p style="margin:0;font-size:14px;letter-spacing:3px;color:#00ff88;
                               text-transform:uppercase;font-weight:700;">STUDENT CREDENTIALS</p>
                  </div>
                  <table width="100%" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td width="50%" style="padding:16px 20px;border-right:1px solid #2a2a4a;
                                             border-bottom:1px solid #2a2a4a;">
                        <p style="margin:0;font-size:14px;letter-spacing:2px;color:#a8b2d1;
                                   text-transform:uppercase;margin-bottom:5px;">FULL NAME</p>
                        <p style="margin:0;font-size:20px;font-weight:700;color:#ffffff;">{name}</p>
                      </td>
                      <td width="50%" style="padding:16px 20px;border-bottom:1px solid #2a2a4a;">
                        <p style="margin:0;font-size:14px;letter-spacing:2px;color:#a8b2d1;
                                   text-transform:uppercase;margin-bottom:5px;">ROLL NUMBER</p>
                        <p style="margin:0;font-size:20px;font-weight:700;color:#00ff88;
                                   letter-spacing:1px;">{roll}</p>
                      </td>
                    </tr>
                    <tr>
                      <td colspan="2" style="padding:16px 20px;">
                        <p style="margin:0;font-size:14px;letter-spacing:2px;color:#a8b2d1;
                                   text-transform:uppercase;margin-bottom:6px;">ACCESS LEVEL</p>
                        <span style="display:inline-block;background:#00ff8815;
                                     border:1px solid #00ff8840;border-radius:6px;
                                     padding:5px 16px;font-size:16px;font-weight:700;
                                     color:#00ff88;letter-spacing:2px;">
                          &#9679; FULL ACCESS GRANTED
                        </span>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:0 40px 32px 40px;text-align:center;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background:linear-gradient(135deg,#0a1a0a,#0a0a18,#0a1020);
                          border:1px solid #00ff8830;border-radius:12px;overflow:hidden;">
              <tr>
                <td style="padding:28px 24px;text-align:center;">
                  <p style="margin:0 0 6px 0;font-size:16px;letter-spacing:4px;
                             color:#00ff88;text-transform:uppercase;font-weight:700;">
                    SCAN TO ENTER
                  </p>
                  <p style="margin:0 0 20px 0;font-size:14px;letter-spacing:2px;color:#5a5a8a;">
                    &#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;
                  </p>
                  <div style="display:inline-block;background:#0a0a0f;
                               border:2px solid #00ff88;border-radius:12px;
                               padding:12px;box-shadow:0 0 30px rgba(0,255,136,0.2);">
                    <img src="cid:qr_image" alt="Your QR Entry Code"
                         width="220" height="220"
                         style="display:block;border-radius:6px;"/>
                  </div>
                  <p style="margin:18px 0 4px 0;font-size:14px;letter-spacing:2px;color:#5a5a8a;">
                    &#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;&#9472;
                  </p>
                  <p style="margin:0;font-size:16px;color:#a8b2d1;letter-spacing:1px;">
                    UNIQUE ID: <span style="color:#00c8ff;font-weight:700;">{roll}</span>
                  </p>
                  <p style="margin:6px 0 0 0;font-size:15px;color:#ff6666;">
                    &#9888; Do not share &#8212; one-time use per student
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#0c0c1a;border:1px solid #2a2a4a;
                          border-radius:10px;overflow:hidden;">
              <tr>
                <td style="padding:20px 24px;">
                  <table width="100%" cellpadding="0" cellspacing="0">
                    <tr><td style="padding:6px 0;">
                      <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                        <span style="color:#00ff88;font-weight:700;">&#8594;</span>
                        Bring this QR code on your phone or as a printout.
                      </p>
                    </td></tr>
                    <tr><td style="padding:6px 0;">
                      <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                        <span style="color:#00ff88;font-weight:700;">&#8594;</span>
                        Gates open <strong style="color:#ffffff;">30 minutes</strong> before event start.
                      </p>
                    </td></tr>
                    <tr><td style="padding:6px 0;">
                      <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                        <span style="color:#00ff88;font-weight:700;">&#8594;</span>
                        Your university ID card must match your registration.
                      </p>
                    </td></tr>
                    <tr><td style="padding:6px 0;">
                      <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                        <span style="color:#00ff88;font-weight:700;">&#8594;</span>
                        This pass is valid for <strong style="color:#ffffff;">both days</strong>: 6 &amp; 7 May 2026.
                      </p>
                    </td></tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td width="33%" style="padding:0 6px 0 0;text-align:center;">
                  <a href="https://github.com/HamzaAhmad-098/secure-entry-qr-system"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">&#128187;</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00ff88;
                               text-transform:uppercase;font-weight:700;">SYSTEM</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">GitHub Repo</p>
                  </a>
                </td>
                <td width="33%" style="padding:0 3px;text-align:center;">
                  <a href="https://chat.whatsapp.com/Lp5J6wNG3Ep99gAmfDSbAp"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">&#128172;</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00ff88;
                               text-transform:uppercase;font-weight:700;">COMMUNITY</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">WhatsApp Group</p>
                  </a>
                </td>
                <td width="33%" style="padding:0 0 0 6px;text-align:center;">
                  <a href="https://maps.google.com/?q=University+of+Engineering+and+Technology+Lahore"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">&#128205;</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00c8ff;
                               text-transform:uppercase;font-weight:700;">VENUE</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">Get Directions</p>
                  </a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:0;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#080810;border-top:1px solid #1a1a2e;">
              <tr>
                <td style="padding:28px 40px;text-align:center;">
                  <div style="height:1px;background:linear-gradient(90deg,transparent,#00ff8840,#0088ff40,transparent);margin-bottom:20px;"></div>
                  <p style="margin:0 0 4px 0;font-size:18px;font-weight:700;
                             color:#ffffff;letter-spacing:2px;text-transform:uppercase;">
                    Welcome Fest 2025
                  </p>
                  <p style="margin:0 0 12px 0;font-size:15px;color:#a8b2d1;letter-spacing:1px;">
                    University of Engineering &amp; Technology, Lahore
                  </p>
                  <p style="margin:0 0 16px 0;font-size:15px;color:#5a5a8a;">
                    Entry System designed &amp; developed by
                    <span style="color:#00ff88;font-weight:700;">Hamza Ahmad CS-12</span>
                    Session 2024
                  </p>
                  <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="text-align:center;">
                        <a href="https://github.com/HamzaAhmad-098/secure-entry-qr-system"
                           style="display:inline-block;margin:0 6px;font-size:15px;
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">GitHub</a>
                        <span style="color:#2a2a4a;">|</span>
                        <a href="https://chat.whatsapp.com/Lp5J6wNG3Ep99gAmfDSbAp"
                           style="display:inline-block;margin:0 6px;font-size:15px;
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">WhatsApp</a>
                        <span style="color:#2a2a4a;">|</span>
                        <a href="https://maps.google.com/?q=University+of+Engineering+and+Technology+Lahore"
                           style="display:inline-block;margin:0 6px;font-size:15px;
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">Venue Map</a>
                      </td>
                    </tr>
                  </table>
                  <div style="height:1px;background:linear-gradient(90deg,transparent,#0088ff40,#00ff8840,transparent);margin-top:20px;margin-bottom:14px;"></div>
                  <p style="margin:0;font-size:14px;color:#ff8888;letter-spacing:2px;text-transform:uppercase;">
                    This QR code is unique to you &#8212; do not forward or share.
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>"""

    msg.attach(MIMEText(html, "html"))
    img_part = MIMEImage(qr_bytes, _subtype="png")
    img_part.add_header("Content-ID", "<qr_image>")
    img_part.add_header("Content-Disposition", "inline", filename=f"QR_{{roll}}.png")
    msg.attach(img_part)
    return msg

def fetch_bounced_emails() -> set:
    """
    Calls the Brevo API to fetch all hard-bounced email addresses
    from the last BOUNCE_LOOKBACK_DAYS days.
    Returns a set of lowercase bounced email addresses.
    Requires BREVO_API_KEY in your .env file.
    """
    if not BREVO_API_KEY:
        log.warning("BREVO_API_KEY not set — skipping bounce detection.")
        log.warning("Add BREVO_API_KEY to your .env to enable this feature.")
        return set()

    bounced = set()
    from datetime import timezone
    now        = datetime.now(timezone.utc)
    start_date = (now - timedelta(days=BOUNCE_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    end_date   = now.strftime("%Y-%m-%d")
    headers = {"api-key": BREVO_API_KEY, "Content-Type": "application/json"}
    offset  = 0
    limit   = 100   # Brevo max per page

    log.info(f"Fetching bounces from Brevo API (last {BOUNCE_LOOKBACK_DAYS} days)...")

    while True:
        try:
            resp = requests.get(
                "https://api.brevo.com/v3/smtp/statistics/events",
                headers=headers,
                params={
                    "event":     "hardBounces",
                    "startDate": start_date,
                    "endDate":   end_date,
                    "limit":     limit,
                    "offset":    offset,
                },
                timeout=15
            )

            if resp.status_code == 401:
                log.error("Brevo API key invalid or expired. Check BREVO_API_KEY in .env")
                break
            if resp.status_code == 403:
                log.error("Brevo API key does not have permission. Enable 'Email Campaigns' in API key settings.")
                break
            if resp.status_code != 200:
                log.error(f"Brevo API error {resp.status_code}: {resp.text}")
                break

            data   = resp.json()
            events = data.get("events", [])

            if not events:
                break   # no more pages

            for event in events:
                email_addr = event.get("email", "").strip().lower()
                if email_addr:
                    bounced.add(email_addr)

            offset += limit
            if offset >= data.get("count", 0):
                break   # fetched all pages

            time.sleep(0.3)   # be polite to the API

        except requests.exceptions.Timeout:
            log.error("Brevo API timed out during bounce fetch.")
            break
        except requests.exceptions.ConnectionError:
            log.error("No internet connection — cannot fetch bounces.")
            break
        except Exception as e:
            log.error(f"Unexpected error fetching bounces: {e}")
            log.debug(traceback.format_exc())
            break

    log.info(f"Bounce detection complete — {len(bounced)} hard bounces found.")
    return bounced


def save_bounced_csv(bounced_emails: set, df: pd.DataFrame):
    """
    Matches bounced email addresses against the student Excel sheet
    and writes bounced_recipients.csv with full student details.
    Also removes bounced emails from sent_log.txt so they can be
    retried with a corrected address if needed.
    """
    if not bounced_emails:
        log.info("No bounced emails to record.")
        return

    # Match bounced emails to student rows
    bounced_rows = []
    for _, row in df.iterrows():
        addr = str(row.get(COL_EMAIL, "")).strip().lower()
        if addr in bounced_emails:
            bounced_rows.append({
                "Student Name": str(row.get(COL_NAME, "")).strip(),
                "Roll Number":  str(row.get(COL_ROLL, "")).strip(),
                "Email":        addr,
                "Status":       "Hard Bounce",
                "Reason":       "Email bounced — address may be invalid or inbox full"
            })

    if not bounced_rows:
        log.info("No bounced emails matched students in the Excel sheet.")
        return

    # Write bounced_recipients.csv
    try:
        with open(BOUNCED_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["Student Name","Roll Number","Email","Status","Reason"])
            w.writeheader()
            w.writerows(bounced_rows)
        log.info(f"Bounced recipients saved → '{BOUNCED_FILE}' ({len(bounced_rows)} entries)")
    except Exception as e:
        log.error(f"Could not write bounced CSV: {e}")

    # Remove bounced emails from sent_log so admin can fix & resend
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, encoding="utf-8") as f:
            current = {line.strip().lower() for line in f if line.strip()}
        cleaned = current - bounced_emails
        removed = len(current) - len(cleaned)
        if removed > 0:
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                for addr in sorted(cleaned):
                    f.write(addr + "\n")
            log.info(f"Removed {removed} bounced address(es) from sent_log.txt")
            log.info("Fix the email addresses in Excel and re-run to retry those students.")


def validate_row(row) -> tuple:
    name  = str(row.get(COL_NAME,  "")).strip()
    roll  = str(row.get(COL_ROLL,  "")).strip()
    email = str(row.get(COL_EMAIL, "")).strip()
    if not name  or name.lower()  == "nan": return False, "Missing name"
    if not roll  or roll.lower()  == "nan": return False, "Missing roll number"
    if "@" not in email or "." not in email.split("@")[-1]: return False, f"Invalid email: {email}"
    return True, ""

def connect_smtp():
    s = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
    s.ehlo()
    s.starttls()
    s.login(SMTP_LOGIN, SENDER_PASS)  # login with Brevo account, not sender
    return s

def main():
    log.info("=" * 55)
    log.info("  Welcome Fest 2025 — QR Email Sender (Brevo)")
    log.info(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 55)

    if not SENDER_PASS:
        log.error("BREVO_SMTP_KEY is not set!")
        log.error("Create a .env file with: BREVO_SMTP_KEY=your-key-here")
        return

    if not os.path.exists(EXCEL_FILE):
        log.error(f"File not found: '{EXCEL_FILE}'")
        return

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
        df.columns = df.columns.str.strip()
        df = df.dropna(how="all")
    except Exception as e:
        log.error(f"Cannot read Excel: {e}"); return

    missing = {COL_NAME, COL_ROLL, COL_EMAIL} - set(df.columns)
    if missing:
        log.error(f"Missing columns: {missing}")
        log.error(f"Found columns:   {list(df.columns)}"); return

    sent_log = load_sent_log()
    total    = len(df)
    pending  = sum(1 for _, r in df.iterrows()
                   if str(r.get(COL_EMAIL,"")).strip().lower() not in sent_log)

    log.info(f"Total rows    : {total}")
    log.info(f"Already sent  : {len(sent_log)}")
    log.info(f"Pending       : {pending}")
    log.info(f"Sending now   : {min(pending, DAILY_LIMIT)} (limit: {DAILY_LIMIT}/day)")
    log.info("-" * 55)

    if pending == 0:
        log.info("All emails already sent! Nothing to do."); return

    try:
        server = connect_smtp()
        log.info("Brevo SMTP connected ✓")
    except smtplib.SMTPAuthenticationError:
        log.error("Authentication failed — check your BREVO_EMAIL and BREVO_SMTP_KEY")
        return
    except Exception as e:
        log.error(f"Connection failed: {e}"); return

    sent_ok = skipped = failed = sent_this_run = 0
    failed_entries = []

    for idx, row in df.iterrows():
        row_num = idx + 2

        if sent_this_run >= DAILY_LIMIT:
            log.warning(f"Daily limit ({DAILY_LIMIT}) reached. Run again tomorrow for the rest.")
            break

        valid, reason = validate_row(row)
        if not valid:
            log.warning(f"Row {row_num}: {reason} — skipped")
            failed_entries.append({"Student Name": str(row.get(COL_NAME,"")).strip(),
                                    "Roll Number": str(row.get(COL_ROLL,"")).strip(),
                                    "Email": str(row.get(COL_EMAIL,"")).strip(),
                                    "Status": "Skipped", "Reason": reason})
            skipped += 1; continue

        name  = str(row[COL_NAME]).strip()
        roll  = str(row[COL_ROLL]).strip()
        email = str(row[COL_EMAIL]).strip().lower()

        if email in sent_log:
            log.info(f"[{row_num}/{total}] SKIP: {email}")
            skipped += 1; continue

        try:
            qr_bytes = generate_qr_bytes(f"Name:{name} | Roll:{roll} | Email:{email}")
            msg      = build_email(email, name, roll, qr_bytes)

            try:
                server.noop()
            except Exception:
                log.warning("Reconnecting SMTP...")
                server = connect_smtp()

            server.sendmail(SENDER_EMAIL, email, msg.as_string())
            mark_sent(email)
            sent_log.add(email)
            sent_ok += 1
            sent_this_run += 1
            log.info(f"[{row_num}/{total}] ✓ SENT → {name} <{email}>")

        except smtplib.SMTPRecipientsRefused:
            log.error(f"[{row_num}/{total}] ✗ REFUSED: {email}")
            failed_entries.append({"Student Name":name,"Roll Number":roll,"Email":email,"Status":"Failed","Reason":"Recipient refused"})
            failed += 1
        except Exception as e:
            log.error(f"[{row_num}/{total}] ✗ ERROR {email}: {e}")
            log.debug(traceback.format_exc())
            failed_entries.append({"Student Name":name,"Roll Number":roll,"Email":email,"Status":"Failed","Reason":str(e)})
            failed += 1

        time.sleep(DELAY_SECONDS)

    try: server.quit()
    except Exception: pass

    if failed_entries:
        with open(FAILED_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["Student Name","Roll Number","Email","Status","Reason"])
            w.writeheader(); w.writerows(failed_entries)
        log.info(f"Failed entries saved → '{FAILED_FILE}'")

    remaining = max(0, pending - sent_this_run)
    log.info("=" * 55)
    log.info(f"  Sent    : {sent_ok}")
    log.info(f"  Skipped : {skipped}")
    log.info(f"  Failed  : {failed}")
    log.info(f"  Left    : {remaining} (run again tomorrow if > 0)")
    log.info("=" * 55)

    # ── BOUNCE DETECTION ──────────────────────────────────────────
    # Runs after every session. Checks Brevo for hard bounces and
    # writes bounced_recipients.csv with full student details.
    log.info("")
    log.info("Running bounce detection via Brevo API...")
    bounced = fetch_bounced_emails()
    save_bounced_csv(bounced, df)
    log.info(f"All done. Check '{BOUNCED_FILE}' for bounced addresses.")

if __name__ == "__main__":
    main()