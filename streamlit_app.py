import io
import re
import time
import html
from collections import defaultdict
from email.header import Header
from email.utils import formataddr, parseaddr

import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------- CONFIG ----------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
USE_TLS = True

EMAILS_PER_BATCH = 5
BATCH_COOLDOWN = 120
DELAY_BETWEEN_EMAILS = 22

# ---------------- PAGE ----------------
st.set_page_config(page_title="Team Niwrutti")
st.title("📧 Team Niwrutti – Smart Bulk Mailer")

# ---------------- SESSION STATE ----------------
defaults = {
    "stop_sending": False,
    "sent_count": 0,
    "last_sent_index": -1,
    "resume_mode": False,
    "failed_rows": []
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------- HELPERS ----------------
def clean_value(val):
    if isinstance(val, str):
        return val.replace("\xa0", " ").replace("\u200b", "").strip()
    return val

def clean_email_address(raw_email):
    if not raw_email:
        return None
    raw_email = clean_value(raw_email)
    _, addr = parseaddr(raw_email)
    return addr if "@" in addr else None

def safe_format(template, mapping):
    return template.format_map(defaultdict(str, mapping))

def get_first_name(full_name: str) -> str:
    if not full_name:
        return ""
    return full_name.strip().split()[0]

def text_to_html(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = html.escape(text)

    paragraphs = text.split("\n\n")
    html_blocks = []

    for para in paragraphs:
        para = para.replace("\n", "<br>")
        html_blocks.append(
            f"<p style='margin:0 0 16px 0; line-height:1.6;'>{para}</p>"
        )

    return "".join(html_blocks)

# ---------------- CSV UPLOAD ----------------
st.subheader("Upload Recipient CSV")
uploaded_file = st.file_uploader("CSV file", type=["csv"])

df = None
if uploaded_file:
    df = pd.read_csv(uploaded_file).applymap(clean_value)
    st.success("CSV uploaded successfully")
    st.dataframe(df)

# ---------------- EMAIL CONFIG ----------------
st.subheader("Email Configuration")
from_email = st.text_input("Gmail address")
app_password = st.text_input("App password", type="password")
from_name = st.text_input("Sender name")

# ---------------- MESSAGE ----------------
st.subheader("Compose Email")
subject_tpl = st.text_input("Subject")
body_tpl = st.text_area("Body", height=450)

# ---------------- STATUS ----------------
st.metric("Emails Sent", st.session_state.sent_count)
progress = st.progress(0)

# ---------------- BUTTONS ----------------
c1, c2, c3 = st.columns(3)
send_btn = c1.button("▶ Send")
stop_btn = c2.button("⛔ Stop")
resume_btn = c3.button("🔁 Resume")

if stop_btn:
    st.session_state.stop_sending = True

if send_btn:
    st.session_state.sent_count = 0
    st.session_state.last_sent_index = -1
    st.session_state.failed_rows = []
    st.session_state.resume_mode = False
    st.session_state.stop_sending = False

if resume_btn:
    st.session_state.resume_mode = True
    st.session_state.stop_sending = False

# ---------------- SEND LOGIC ----------------
def send_bulk(df_to_send, resume=False):
    start_index = (
        st.session_state.last_sent_index + 1
        if resume else 0
    )

    total = len(df_to_send)
    sent_in_batch = 0

    for idx in range(start_index, total):

        if st.session_state.stop_sending:
            st.warning("Sending stopped. You can resume later.")
            break

        row = df_to_send.iloc[idx].to_dict()
        recip = clean_email_address(row.get("email"))
        if not recip:
            continue

        subject = safe_format(subject_tpl, row)

        body_row = dict(row)
        body_row["name"] = get_first_name(row.get("name", ""))

        body_text = safe_format(body_tpl, body_row)

        html_body = f"""
        <html>
          <body style="font-family:'Times New Roman', serif; font-size:15px;">
            {text_to_html(body_text)}
          </body>
        </html>
        """

        msg = MIMEMultipart()
        msg["From"] = formataddr((str(Header(from_name, "utf-8")), from_email))
        msg["To"] = recip
        msg["Subject"] = str(Header(subject, "utf-8"))
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(from_email, app_password)
                server.send_message(msg)

            st.success(f"Sent to {recip}")
            st.session_state.sent_count += 1
            st.session_state.last_sent_index = idx
            sent_in_batch += 1

        except Exception as e:
            st.error(f"Failed: {recip}")
            st.session_state.failed_rows.append({**row, "__error": str(e)})

        progress.progress((idx + 1) / total)
        time.sleep(DELAY_BETWEEN_EMAILS)

        if sent_in_batch >= EMAILS_PER_BATCH:
            st.warning("Cooling down for Gmail safety...")
            time.sleep(BATCH_COOLDOWN)
            sent_in_batch = 0

# ---------------- EXECUTION ----------------
if (send_btn or resume_btn) and df is not None:
    send_bulk(df, resume=st.session_state.resume_mode)

# ---------------- RETRY FAILED ----------------
if st.session_state.failed_rows:
    st.subheader("❌ Failed Emails")
    failed_df = pd.DataFrame(st.session_state.failed_rows)
    st.dataframe(failed_df)

    if st.button("🔄 Retry Failed Emails Only"):
        st.session_state.stop_sending = False
        st.session_state.last_sent_index = -1
        send_bulk(failed_df, resume=False)

# ---------------- FOOTER ----------------
st.markdown(
    f"**Total Sent:** {st.session_state.sent_count} | "
    f"**Failed:** {len(st.session_state.failed_rows)}"
)
