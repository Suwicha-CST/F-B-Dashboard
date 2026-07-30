import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASS = os.environ["SMTP_PASS"]
RECIPIENTS = [r.strip() for r in os.environ["RECIPIENTS"].split(",")]

DASHBOARD_URL = "https://suwicha-cst.github.io/F-B-Dashboard/"

msg = MIMEMultipart("related")
msg["Subject"] = "Jul's & Zephyr — Daily Performance Snapshot"
msg["From"] = SMTP_USER
msg["To"] = ", ".join(RECIPIENTS)

html = f"""
<html>
  <body style="font-family: Arial, sans-serif;">
    <p>Good morning,</p>
    <p>Here's today's snapshot of the Jul's &amp; Zephyr dashboard:<br>
       For the live, interactive version where you can filter by outlet, date, or period:
       <a href="{DASHBOARD_URL}">click here</a></p>
    <img src="cid:dashboard_image" style="max-width:700px; width:100%; border:1px solid #ddd;" />
  </body>
</html>
"""
msg.attach(MIMEText(html, "html"))

with open("dashboard.png", "rb") as f:
    img = MIMEImage(f.read())
    img.add_header("Content-ID", "<dashboard_image>")
    img.add_header("Content-Disposition", "inline", filename="dashboard.png")
    msg.attach(img)

with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
    server.starttls()
    server.login(SMTP_USER, SMTP_PASS)
    server.sendmail(SMTP_USER, RECIPIENTS, msg.as_string())

print("Email sent successfully to:", RECIPIENTS)
