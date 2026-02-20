import os
import smtplib
from email.message import EmailMessage

from flask import Flask, render_template, request, abort
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
MAIL_FROM = os.getenv("MAIL_FROM", SMTP_USER)

# 간단한 보호용 토큰(없어도 되지만, 외부 노출 시 무조건 두세요)
SEND_TOKEN = os.getenv("SEND_TOKEN", "")

def send_email(to_addr: str, subject: str, body: str) -> None:
    if not SMTP_HOST or not SMTP_USER or not SMTP_PASS or not MAIL_FROM:
        raise RuntimeError("SMTP 설정(SMTP_HOST/SMTP_USER/SMTP_PASS/MAIL_FROM)이 비어 있습니다.")

    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)

    # STARTTLS(587) 기준
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as smtp:
        smtp.ehlo()
        if SMTP_PORT == 587:
            smtp.starttls()
            smtp.ehlo()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)

@app.get("/")
def index():
    return render_template("index.html")

@app.post("/send")
def send():
    # 외부에 배포할 거면 최소한 토큰 체크 같은 보호장치를 꼭 두세요.
    if SEND_TOKEN:
        token = request.headers.get("X-SEND-TOKEN", "")
        if token != SEND_TOKEN:
            abort(403)

    to_addr = request.form.get("to", "").strip()
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()

    if not to_addr or not subject or not body:
        abort(400)

    send_email(to_addr, subject, body)
    return render_template("sent.html")

if __name__ == "__main__":
    # 로컬 개발용. 운영 배포는 gunicorn/uwsgi 같은 WSGI로 하세요.
    app.run(host="127.0.0.1", port=5000, debug=True)
