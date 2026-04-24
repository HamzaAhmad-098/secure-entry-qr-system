"""
send_qr_emails.py
=================
Reads an Excel file with columns:
    Student Name | Roll Number | Email | QR Code (IMAGE formula or URL text)

Then generates a QR code for each student (from Roll Number)
and sends it to their email as an attachment.

NEW: IMAP bounce detection – after sending, connects to Gmail inbox,
finds undelivered mail reports, removes bounced addresses from sent_log.txt,
and records them as failed.

SETUP:
------
1. pip install pandas openpyxl qrcode pillow python-dotenv
2. Fill in your Gmail credentials (App Password) – same for SMTP and IMAP.
3. Set EXCEL_FILE path.
4. Run:  python send_qr_emails.py
"""

import os
import io
import time
import csv
import logging
import smtplib
import imaplib
import email
import ssl
import re
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import decode_header
from email.utils import parseaddr

import pandas as pd
import qrcode
from dotenv import load_dotenv
load_dotenv()

# ──────────────────────────────────────────────
#  CONFIG — Edit these before running
# ──────────────────────────────────────────────
EXCEL_FILE   = "students.xlsx"
SHEET_NAME   = 0

COL_NAME     = "Student Name"
COL_ROLL     = "Roll Number"
COL_EMAIL    = "Email"

SENDER_EMAIL = "hamzaxdevelopers1223@gmail.com"
SENDER_PASS  = os.getenv("appPassword")       # Gmail App Password (16 chars)
SMTP_HOST    = "smtp.gmail.com"
SMTP_PORT    = 587

# IMAP settings (for bounce detection)
IMAP_HOST    = "imap.gmail.com"
IMAP_PORT    = 993
IMAP_USER    = SENDER_EMAIL                  # same as sender
IMAP_PASS    = SENDER_PASS                   # same App Password

SUBJECT      = "Your Student QR Code"

DELAY_SECONDS = 1.5
LOG_FILE     = "sent_log.txt"
FAILED_LOG_FILE = "failed_recipients.csv"    # students who could NOT be sent
BOUNCED_LOG_FILE = "bounced_recipients.csv"  # students whose email bounced after sending
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


def rewrite_sent_log(emails: set):
    """Overwrite sent_log.txt with a fresh set of email addresses."""
    with open(LOG_FILE, "w") as f:
        for em in sorted(emails):
            f.write(em + "\n")


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

    html_body = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Welcome Fest 2025 — Entry Pass</title>
</head>
<body style="margin:0;padding:0;background-color:#05050a;font-family:'Segoe UI',Arial,sans-serif;">
<!-- ░░ WRAPPER ░░ -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="background-color:#05050a;min-height:100vh;">
  <tr>
    <td align="center" style="padding:40px 16px;">
 
      <!-- ░░ CARD ░░ -->
      <table width="650" cellpadding="0" cellspacing="0" border="0"
             style="max-width:650px;width:100%;background-color:#0d0d16;
                    border:1px solid #1a1a2e;border-radius:16px;
                    overflow:hidden;box-shadow:0 0 60px rgba(0,255,136,0.08);">
 
        <!-- ══ HERO HEADER ══ -->
        <tr>
          <td style="padding:0;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background:linear-gradient(135deg,#0a0a14 0%,#0f1628 50%,#0a0a14 100%);
                          border-bottom:1px solid #1a2a1a;">
              <tr>
                <td style="padding:36px 40px 28px 40px;text-align:center;">
 
                  <div style="margin-bottom:20px;">
                    <div style="height:2px;background:linear-gradient(90deg,transparent,#00ff88,#0088ff,transparent);margin-bottom:24px;"></div>
                  </div>
 
                  <p style="margin:0 0 6px 0;font-size:15px;letter-spacing:4px;
                             color:#00ff88;text-transform:uppercase;font-weight:700;">
                    UNIVERSITY OF ENGINEERING &amp; TECHNOLOGY, LAHORE
                  </p>
                  <p style="margin:0 0 18px 0;font-size:14px;letter-spacing:2px;
                             color:#7fffaa;text-transform:uppercase;">
                    SESSION 2024 PRESENTS
                  </p>
 
                  <h1 style="margin:0;font-size:52px;font-weight:900;letter-spacing:-1px;
                              line-height:1.1;">
                    <span style="color:#ffffff;">WELCOME</span>
                    <span style="color:#00ff88;"> FEST</span>
                    <br/>
                    <span style="color:#00c8ff;font-size:58px;letter-spacing:2px;">2025</span>
                  </h1>
 
                  <p style="margin:14px 0 0 0;font-size:16px;letter-spacing:3px;
                             color:#c0caf5;text-transform:uppercase;">
                    &lt; YOUR OFFICIAL ENTRY PASS /&gt;
                  </p>
 
                  <div style="height:2px;background:linear-gradient(90deg,transparent,#0088ff,#00ff88,transparent);margin-top:24px;"></div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
        <!-- ══ DATE & VENUE STRIP ══ -->
        <tr>
          <td style="padding:0;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#080810;border-bottom:1px solid #1a1a2e;">
              <tr>
                <td width="50%" style="padding:18px 24px;border-right:1px solid #1a1a2e;text-align:center;">
                  <p style="margin:0;font-size:14px;letter-spacing:3px;color:#00ff88;text-transform:uppercase;margin-bottom:4px;">
                    📅 DATE
                  </p>
                  <p style="margin:0;font-size:20px;font-weight:700;color:#ffffff;letter-spacing:1px;">
                    6 &amp; 7 MAY 2026
                  </p>
                </td>
                <td width="50%" style="padding:18px 24px;text-align:center;">
                  <p style="margin:0;font-size:14px;letter-spacing:3px;color:#00c8ff;text-transform:uppercase;margin-bottom:4px;">
                    📍 VENUE
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
 
        <!-- ══ GREETING ══ -->
        <tr>
          <td style="padding:32px 40px 20px 40px;">
            <p style="margin:0;font-size:16px;letter-spacing:3px;color:#00ff88;
                      text-transform:uppercase;margin-bottom:8px;">
            </p>
            <h2 style="margin:0;font-size:36px;font-weight:800;color:#ffffff;
                       line-height:1.2;">
              Hello, {name}!
            </h2>
            <p style="margin:14px 0 0 0;font-size:18px;color:#ccd6f6;line-height:1.7;">
              Your <span style="color:#00ff88;font-weight:700;">personal entry pass</span>
              for Welcome Fest 2025 has been generated and is ready to use.
              Present this QR code at the entrance gate for instant verification.
            </p>
          </td>
        </tr>
 
        <!-- ══ STUDENT INFO CARD ══ -->
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#0a0a18;border:1px solid #2a2a4a;
                          border-radius:10px;overflow:hidden;">
              <tr>
                <td style="padding:0;">
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
                        <p style="margin:0;">
                          <span style="display:inline-block;background:#00ff8815;
                                       border:1px solid #00ff8840;border-radius:6px;
                                       padding:5px 16px;font-size:16px;font-weight:700;
                                       color:#00ff88;letter-spacing:2px;">
                            ● FULL ACCESS GRANTED
                          </span>
                        </p>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
        <!-- ══ QR CODE SECTION ══ -->
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
                    ─────────────────────────────
                  </p>
 
                  <div style="display:inline-block;background:#0a0a0f;
                               border:2px solid #00ff88;border-radius:12px;
                               padding:12px;box-shadow:0 0 30px rgba(0,255,136,0.2);">
                    <img src="cid:qr_image" alt="Your QR Entry Code"
                         width="220" height="220"
                         style="display:block;border-radius:6px;"/>
                  </div>
 
                  <p style="margin:18px 0 4px 0;font-size:14px;letter-spacing:2px;color:#5a5a8a;">
                    ─────────────────────────────
                  </p>
                  <p style="margin:0;font-size:16px;color:#a8b2d1;letter-spacing:1px;">
                    UNIQUE ID: <span style="color:#00c8ff;font-weight:700;">{roll}</span>
                  </p>
                  <p style="margin:6px 0 0 0;font-size:15px;color:#ff6666;">
                    ⚠ Do not share — one-time use per student
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
        <!-- ══ IMPORTANT INSTRUCTIONS ══ -->
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0"
                   style="background-color:#0c0c1a;border:1px solid #2a2a4a;
                          border-radius:10px;overflow:hidden;">
              <tr>
                <td style="padding:20px 24px;">
                  <p style="margin:0 0 14px 0;font-size:16px;letter-spacing:3px;
                             color:#00c8ff;text-transform:uppercase;font-weight:700;">
                  </p>
                  <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="padding:6px 0;">
                        <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                          <span style="color:#00ff88;font-weight:700;">→</span>
                          Bring this QR code on your phone or as a printout.
                        </p>
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;">
                        <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                          <span style="color:#00ff88;font-weight:700;">→</span>
                          Gates open <strong style="color:#ffffff;">30 minutes</strong> before event start.
                        </p>
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;">
                        <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                          <span style="color:#00ff88;font-weight:700;">→</span>
                          Your university ID card must match your registration.
                        </p>
                      </td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;">
                        <p style="margin:0;font-size:17px;color:#ccd6f6;line-height:1.6;">
                          <span style="color:#00ff88;font-weight:700;">→</span>
                          This pass is valid for <strong style="color:#ffffff;">both days</strong>: 6 &amp; 7 May 2026.
                        </p>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
        <!-- ══ LINKS SECTION ══ -->
        <tr>
          <td style="padding:0 40px 28px 40px;">
            <p style="margin:0 0 14px 0;font-size:16px;letter-spacing:3px;
                       color:#a8b2d1;text-transform:uppercase;text-align:center;">
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td width="33%" style="padding:0 6px 0 0;text-align:center;">
                  <a href="https://github.com/HamzaAhmad-098/secure-entry-qr-system"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">💻</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00ff88;
                               text-transform:uppercase;font-weight:700;">SYSTEM</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">GitHub Repo</p>
                  </a>
                </td>
                <td width="33%" style="padding:0 3px;text-align:center;">
                  <a href="https://chat.whatsapp.com/Lp5J6wNG3Ep99gAmfDSbAp"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">💬</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00ff88;
                               text-transform:uppercase;font-weight:700;">COMMUNITY</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">WhatsApp Group</p>
                  </a>
                </td>
                <td width="33%" style="padding:0 0 0 6px;text-align:center;">
                  <a href="https://maps.google.com/?q=University+of+Engineering+and+Technology+Lahore"
                     style="display:block;background:#0a0a18;border:1px solid #2a2a4a;
                            border-radius:8px;padding:14px 10px;text-decoration:none;">
                    <p style="margin:0 0 6px 0;font-size:22px;">📍</p>
                    <p style="margin:0;font-size:15px;letter-spacing:2px;color:#00c8ff;
                               text-transform:uppercase;font-weight:700;">VENUE</p>
                    <p style="margin:3px 0 0 0;font-size:14px;color:#a8b2d1;">Get Directions</p>
                  </a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
        <!-- ══ FOOTER ══ -->
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
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">
                          GitHub
                        </a>
                        <span style="color:#2a2a4a;">|</span>
                        <a href="https://chat.whatsapp.com/Lp5J6wNG3Ep99gAmfDSbAp"
                           style="display:inline-block;margin:0 6px;font-size:15px;
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">
                          WhatsApp
                        </a>
                        <span style="color:#2a2a4a;">|</span>
                        <a href="https://maps.google.com/?q=University+of+Engineering+and+Technology+Lahore"
                           style="display:inline-block;margin:0 6px;font-size:15px;
                                  color:#ccd6f6;text-decoration:none;letter-spacing:1px;">
                          Venue Map
                        </a>
                      </td>
                    </tr>
                  </table>
 
                  <div style="height:1px;background:linear-gradient(90deg,transparent,#0088ff40,#00ff8840,transparent);margin-top:20px;margin-bottom:14px;"></div>
 
                  <p style="margin:0;font-size:14px;color:#ff8888;letter-spacing:2px;text-transform:uppercase;">
                    This QR code is unique to you — do not forward or share.
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
 
      </table>
      <!-- END CARD -->
 
    </td>
  </tr>
</table>
 
</body>
</html>"""

    msg.attach(MIMEText(html_body, "html"))

    # Embed QR image inline
    img_part = MIMEImage(qr_bytes, _subtype="png")
    img_part.add_header("Content-ID", "<qr_image>")
    img_part.add_header(
        "Content-Disposition", "inline",
        filename=f"QR_{roll}.png"
    )
    msg.attach(img_part)
    return msg


def validate_row(row, idx: int):
    """
    Returns (is_valid: bool, reason: str).
    """
    name  = str(row.get(COL_NAME,  "")).strip()
    roll  = str(row.get(COL_ROLL,  "")).strip()
    email = str(row.get(COL_EMAIL, "")).strip()

    if not name or name.lower() == "nan":
        return False, "Missing student name"
    if not roll or roll.lower() == "nan":
        return False, "Missing roll number"
    if "@" not in email or "." not in email.split("@")[-1]:
        return False, f"Invalid email format: '{email}'"
    return True, ""


def extract_recipient_from_bounce(msg):
    """
    Extract the original recipient email address from a bounce DSN.
    Returns the email string or None.
    """
    # 1. Standard headers that often carry the failed recipient directly
    for header in ["X-Failed-Recipients", "X-Original-To", "X-Envelope-To"]:
        if header in msg:
            val = msg[header].strip()
            if "@" in val:
                return val.lower()

    # 2. Walk all parts for delivery-status sections or plain text bodies
    candidate = None
    for part in msg.walk():
        content_type = part.get_content_type()

        if content_type == "message/delivery-status":
            # delivery-status is a Message object; its headers contain Final-Recipient
            for key in ["Final-Recipient", "Original-Recipient"]:
                val = part.get(key)
                if val:
                    # format: rfc822; address@domain
                    parts = val.split(";", 1)
                    if len(parts) > 1:
                        addr = parts[1].strip()
                        if "@" in addr:
                            return addr.lower()

        elif content_type in ("text/plain", "text/html"):
            try:
                body = part.get_payload(decode=True)
                if body:
                    text = body.decode("utf-8", errors="ignore")
            except Exception:
                continue

            # --- Try many known bounce patterns ---

            # Gmail style: "The following address(es) failed:"
            match = re.search(
                r"(?:following\s+address(?:es)?\s+failed:\s*)(<?[^\s@>]+@[^\s>]+>?)",
                text, re.IGNORECASE)
            if match:
                return match.group(1).strip("<> ").lower()

            # Generic: "could not be delivered to <email>"
            match = re.search(
                r"could\s+not\s+be\s+delivered\s+to\s+<?([\w.\-]+@[\w.\-]+)>?",
                text, re.IGNORECASE)
            if match:
                return match.group(1).lower()

            # "Message was not delivered to <email>"
            match = re.search(
                r"message\s+was\s+not\s+delivered\s+to\s+<?([\w.\-]+@[\w.\-]+)>?",
                text, re.IGNORECASE)
            if match:
                return match.group(1).lower()

            # "Final-Recipient: rfc822; email"
            match = re.search(
                r"Final-Recipient:\s*rfc822;\s*<?([\w.\-]+@[\w.\-]+)>?",
                text, re.IGNORECASE)
            if match:
                return match.group(1).lower()

            # "Original-Recipient: rfc822; email"
            match = re.search(
                r"Original-Recipient:\s*rfc822;\s*<?([\w.\-]+@[\w.\-]+)>?",
                text, re.IGNORECASE)
            if match:
                return match.group(1).lower()

            # Fallback: extract the first email that is not our own sender
            # (only if we haven't found anything better)
            if not candidate:
                emails = re.findall(r"[\w.\-]+@[\w.\-]+", text.lower())
                # exclude the sender address
                for em in emails:
                    if em != SENDER_EMAIL:
                        candidate = em
                        break

    return candidate.lower() if candidate else None


def detect_bounces():
    """
    Connect to IMAP, search for bounce/failure messages,
    extract failed recipients, remove them from sent_log.txt,
    and return a set of bounced email addresses.
    """
    bounced_emails = set()
    try:
        context = ssl.create_default_context()
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=context) as mail:
            mail.login(IMAP_USER, IMAP_PASS)
            # Select inbox; add more folders if needed
            mail.select("INBOX")

            # 1) Search for known bounce subjects AND messages from daemon senders
            #    This covers many variations
            subject_keywords = [
                "Undelivered Mail Returned",
                "Delivery Status Notification",
                "Returned mail",
                "Mail delivery failed",
                "mail delivery failure",
                "Delivery Notification",
                "failure notice",
                "Undeliverable",
                "Returned Mail",
            ]
            # Build an OR chain for subjects
            search_str = ""
            for kw in subject_keywords:
                if search_str:
                    search_str = f"OR (SUBJECT \"{kw}\") {search_str}"
                else:
                    search_str = f"(SUBJECT \"{kw}\")"

            # Also search for common bounce from addresses
            search_str = f"OR FROM \"mailer-daemon\" OR FROM \"Mail Delivery System\" {search_str}"

            status, data = mail.search(None, search_str)
            if status != 'OK':
                log.warning("IMAP search did not return OK.")
                return bounced_emails

            mail_ids = data[0].split()
            if not mail_ids:
                log.info("No bounce messages found in inbox.")
                return bounced_emails

            log.info(f"Found {len(mail_ids)} potential bounce messages. Processing...")

            for mail_id in mail_ids:
                status, msg_data = mail.fetch(mail_id, "(RFC822)")
                if status != 'OK':
                    continue
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                recipient = extract_recipient_from_bounce(msg)
                if recipient:
                    bounced_emails.add(recipient)
                    log.info(f"  Bounce detected for: {recipient}")
                else:
                    # Log subject for troubleshooting
                    subject = msg["Subject"] or "(no subject)"
                    log.info(f"  Could not extract recipient from: {subject}")

                # Delete the bounce email after processing
                mail.store(mail_id, '+FLAGS', '\\Deleted')

            mail.expunge()
            log.info(f"Total bounced recipients found: {len(bounced_emails)}")

    except Exception as e:
        log.error(f"IMAP bounce detection failed: {e}")
        log.debug(traceback.format_exc())

    return bounced_emails
def process_bounces(existing_sent_log: set, failed_entries: list, df: pd.DataFrame):
    """
    1. Detect bounce emails.
    2. For bounced addresses, remove from sent_log.txt and add to failed CSV.
    3. Update sent_log.txt file.
    """
    bounced = detect_bounces()
    if not bounced:
        return

    # Remove bounced emails from sent_log set and text file
    affected = existing_sent_log.intersection(bounced)
    if affected:
        log.info(f"Removing from sent_log: {affected}")
        clean_sent = existing_sent_log - bounced
        rewrite_sent_log(clean_sent)

    # Now find the student details from the original Excel to add to failure log
    # (we might have multiple rows with same email, but we'll add all)
    for _, row in df.iterrows():
        email_val = str(row.get(COL_EMAIL, "")).strip().lower()
        if email_val in bounced:
            name = str(row.get(COL_NAME, "")).strip()
            roll = str(row.get(COL_ROLL, "")).strip()
            # Only add if not already in failed_entries (avoid duplicates)
            if not any(entry["Email"] == email_val for entry in failed_entries):
                failed_entries.append({
                    "Student Name": name,
                    "Roll Number": roll,
                    "Email": email_val,
                    "Status": "Failed - Bounced",
                    "Reason": "Email bounced after sending"
                })

    # Write the updated failed_recipients.csv (merge with existing entries)
    if failed_entries:
        try:
            # We'll write all failures (including prior send errors + bounces)
            with open(FAILED_LOG_FILE, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["Student Name", "Roll Number", "Email", "Status", "Reason"])
                writer.writeheader()
                writer.writerows(failed_entries)
            log.info(f"Updated '{FAILED_LOG_FILE}' with {len(failed_entries)} entries.")
        except Exception as e:
            log.error(f"Failed to write failure CSV: {e}")

    # Also write a separate file only for bounced addresses (optional clarity)
    bounced_only = [e for e in failed_entries if e["Status"] == "Failed - Bounced"]
    if bounced_only:
        try:
            with open(BOUNCED_LOG_FILE, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["Student Name", "Roll Number", "Email", "Status", "Reason"])
                writer.writeheader()
                writer.writerows(bounced_only)
            log.info(f"Bounced recipients also saved to '{BOUNCED_LOG_FILE}'.")
        except Exception as e:
            log.error(f"Failed to write bounced CSV: {e}")


def main():
    # ── Load Excel ──────────────────────────────
    if not os.path.exists(EXCEL_FILE):
        log.error(f"Excel file not found: {EXCEL_FILE}")
        return

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
        df.columns = df.columns.str.strip()
        df = df.dropna(how="all")
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
    failed_entries = []   # list of dicts

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
        row_num = idx + 2

        valid, reason = validate_row(row, row_num)
        if not valid:
            log.warning(f"Row {row_num}: {reason} — skipped")
            name_val = str(row.get(COL_NAME, "")).strip()
            roll_val = str(row.get(COL_ROLL, "")).strip()
            email_val = str(row.get(COL_EMAIL, "")).strip()
            failed_entries.append({
                "Student Name": name_val,
                "Roll Number": roll_val,
                "Email": email_val,
                "Status": "Skipped - Invalid Data",
                "Reason": reason
            })
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

            # Reconnect if needed
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
            err_msg = "Recipient refused by server"
            log.error(f"[{row_num}/{total}] {err_msg}: {email}")
            failed_entries.append({
                "Student Name": name,
                "Roll Number": roll,
                "Email": email,
                "Status": "Failed - Recipient Refused",
                "Reason": err_msg
            })
            failed += 1
        except smtplib.SMTPException as e:
            err_msg = str(e)
            log.error(f"[{row_num}/{total}] SMTP error for {email}: {err_msg}")
            failed_entries.append({
                "Student Name": name,
                "Roll Number": roll,
                "Email": email,
                "Status": "Failed - SMTP Error",
                "Reason": err_msg
            })
            failed += 1
        except Exception as e:
            err_msg = str(e)
            log.error(f"[{row_num}/{total}] Unexpected error for {email}: {err_msg}")
            log.debug(traceback.format_exc())
            failed_entries.append({
                "Student Name": name,
                "Roll Number": roll,
                "Email": email,
                "Status": "Failed - Unexpected",
                "Reason": err_msg
            })
            failed += 1

        time.sleep(DELAY_SECONDS)

    # ── Cleanup SMTP connection ─────────────────
    try:
        server.quit()
    except Exception:
        pass

    # ── Write initial failure CSV (pre‑bounce) ──
    if failed_entries:
        try:
            with open(FAILED_LOG_FILE, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["Student Name", "Roll Number", "Email", "Status", "Reason"])
                writer.writeheader()
                writer.writerows(failed_entries)
        except Exception as e:
            log.error(f"Failed to write failure CSV: {e}")

    # ── NOW RUN BOUNCE DETECTION ────────────────
    # Update sent_log (remove bounced) and add bounced entries to failure CSV
    process_bounces(sent_log, failed_entries, df)

    # ── Final Summary ───────────────────────────
    log.info("=" * 50)
    log.info(f"DONE — Total: {total} | Sent: {sent_ok} | Skipped: {skipped} | Failed: {failed}")
    log.info(f"Check 'email_run.log' for full details.")
    if failed:
        log.warning(f"{failed} emails failed. Re-run the script to retry them.")
    log.info(f"Bounce detection completed. See '{BOUNCED_LOG_FILE}' and '{FAILED_LOG_FILE}'.")


if __name__ == "__main__":
    main()