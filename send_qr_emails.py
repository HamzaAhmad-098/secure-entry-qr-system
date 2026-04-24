"""
send_qr_emails.py
=================
Reads an Excel file with columns:
    Student Name | Roll Number | Email | QR Code (IMAGE formula or URL text)

Then generates a QR code for each student (from Roll Number)
and sends it to their email as an attachment.

SETUP:
------
1. pip install pandas openpyxl qrcode pillow

2. Fill in your Gmail credentials in the CONFIG section below.
   For Gmail, use an App Password (not your real password):
   → Google Account → Security → 2-Step Verification → App Passwords
   → Generate one for "Mail"

3. Set EXCEL_FILE to the path of your Excel file.

4. Run:  python send_qr_emails.py
"""

import os
import io
import time
import logging
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

import pandas as pd
import qrcode

# ──────────────────────────────────────────────
#  CONFIG — Edit these before running
# ──────────────────────────────────────────────
EXCEL_FILE   = "students.xlsx"       # Path to your Excel file
SHEET_NAME   = 0                     # Sheet index (0 = first) or name e.g. "Sheet1"

# Column names exactly as they appear in your Excel header row
COL_NAME     = "Student Name"
COL_ROLL     = "Roll Number"
COL_EMAIL    = "Email"

SENDER_EMAIL = "hamzaxdevelopers1223@gmail.com"       # Your Gmail address
SENDER_PASS  =  os.getenv("appPassword")# Gmail App Password (16 chars)
SMTP_HOST    = "smtp.gmail.com"
SMTP_PORT    = 587

SUBJECT      = "Your Student QR Code"

# Delay between emails in seconds (avoid spam filters)
DELAY_SECONDS = 1.5

# Resume support: already-sent emails are logged here
LOG_FILE     = "sent_log.txt"
# ──────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.FileHandler("email_run.log"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)


def load_sent_log():
    """Return set of emails already successfully sent."""
    if not os.path.exists(LOG_FILE):
        return set()
    with open(LOG_FILE) as f:
        return {line.strip() for line in f if line.strip()}


def mark_sent(email: str):
    with open(LOG_FILE, "a") as f:
        f.write(email.strip() + "\n")


def generate_qr_bytes(data: str) -> bytes:
    """Generate a QR code image and return it as PNG bytes."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=3,
    )
    qr.add_data(str(data))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def build_email(sender: str, recipient: str, name: str, roll: str, qr_bytes: bytes) -> MIMEMultipart:
    msg = MIMEMultipart("related")
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = SUBJECT

    html_body = f"""
    <html><body style="font-family:Arial,sans-serif;color:#222;">
      <h2 style="color:#1F4E79;">Hello, {name}!</h2>
      <p>Please find your personal <strong>Student QR Code</strong> attached below.</p>
      <p>
        <strong>Roll Number:</strong> {roll}<br>
        <strong>Email:</strong> {recipient}
      </p>
      <p style="margin-top:20px;">
        <img src="cid:qr_image" alt="QR Code"
             style="border:2px solid #1F4E79;border-radius:8px;padding:6px;" />
      </p>
      <p style="color:#555;font-size:12px;margin-top:30px;">
        This QR code is unique to you. Please do not share it.<br>
        — Student Affairs Office
      </p>
    </body></html>
    """

    msg.attach(MIMEText(html_body, "html"))

    # Embed QR image inline
    img_part = MIMEImage(qr_bytes, _subtype="png")
    img_part.add_header("Content-ID", "<qr_image>")
    img_part.add_header(
        "Content-Disposition", "attachment",
        filename=f"QR_{roll}.png"
    )
    msg.attach(img_part)
    return msg


def validate_row(row, idx: int) -> bool:
    """Return True if the row has valid name, roll, and email."""
    name  = str(row.get(COL_NAME,  "")).strip()
    roll  = str(row.get(COL_ROLL,  "")).strip()
    email = str(row.get(COL_EMAIL, "")).strip()

    if not name or name.lower() == "nan":
        log.warning(f"Row {idx}: missing name — skipped")
        return False
    if not roll or roll.lower() == "nan":
        log.warning(f"Row {idx}: missing roll number — skipped")
        return False
    if "@" not in email or "." not in email.split("@")[-1]:
        log.warning(f"Row {idx}: invalid email '{email}' — skipped")
        return False
    return True


def main():
    # ── Load Excel ──────────────────────────────
    if not os.path.exists(EXCEL_FILE):
        log.error(f"Excel file not found: {EXCEL_FILE}")
        return

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
        df.columns = df.columns.str.strip()          # trim header whitespace
        df = df.dropna(how="all")                    # drop completely blank rows
    except Exception as e:
        log.error(f"Failed to read Excel file: {e}")
        return

    required = {COL_NAME, COL_ROLL, COL_EMAIL}
    missing  = required - set(df.columns)
    if missing:
        log.error(f"Missing columns in Excel: {missing}")
        log.error(f"Found columns: {list(df.columns)}")
        return

    sent_log = load_sent_log()
    total    = len(df)
    sent_ok  = 0
    skipped  = 0
    failed   = 0

    log.info(f"Loaded {total} rows from '{EXCEL_FILE}'")
    log.info(f"Already sent (from log): {len(sent_log)}")

    # ── Connect SMTP ────────────────────────────
    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
        server.ehlo()
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASS)
        log.info("SMTP connection established ✓")
    except smtplib.SMTPAuthenticationError:
        log.error("Authentication failed. Check SENDER_EMAIL and SENDER_PASS.")
        log.error("For Gmail, use an App Password, not your account password.")
        return
    except Exception as e:
        log.error(f"SMTP connection failed: {e}")
        return

    # ── Send loop ───────────────────────────────
    for idx, row in df.iterrows():
        row_num = idx + 2  # human-readable Excel row number

        if not validate_row(row, row_num):
            skipped += 1
            continue

        name  = str(row[COL_NAME]).strip()
        roll  = str(row[COL_ROLL]).strip()
        email = str(row[COL_EMAIL]).strip().lower()

        if email in sent_log:
            log.info(f"[{row_num}/{total}] SKIP (already sent): {email}")
            skipped += 1
            continue

        try:
            qr_bytes = generate_qr_bytes(f"Name:{name} | Roll:{roll} | Email:{email}")
            msg      = build_email(SENDER_EMAIL, email, name, roll, qr_bytes)

            # Reconnect if connection was dropped
            try:
                server.noop()
            except Exception:
                log.warning("SMTP connection dropped, reconnecting...")
                server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
                server.ehlo()
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASS)

            server.sendmail(SENDER_EMAIL, email, msg.as_string())
            mark_sent(email)
            sent_ok += 1
            log.info(f"[{row_num}/{total}] SENT → {name} <{email}>")

        except smtplib.SMTPRecipientsRefused:
            log.error(f"[{row_num}/{total}] REFUSED by server: {email}")
            failed += 1
        except smtplib.SMTPException as e:
            log.error(f"[{row_num}/{total}] SMTP error for {email}: {e}")
            failed += 1
        except Exception as e:
            log.error(f"[{row_num}/{total}] Unexpected error for {email}: {e}")
            log.debug(traceback.format_exc())
            failed += 1

        time.sleep(DELAY_SECONDS)

    # ── Cleanup ─────────────────────────────────
    try:
        server.quit()
    except Exception:
        pass

    # ── Summary ─────────────────────────────────
    log.info("=" * 50)
    log.info(f"DONE — Total: {total} | Sent: {sent_ok} | Skipped: {skipped} | Failed: {failed}")
    log.info(f"Check 'email_run.log' for full details.")
    if failed:
        log.warning(f"{failed} emails failed. Re-run the script to retry them (already-sent are skipped).")


if __name__ == "__main__":
    main()