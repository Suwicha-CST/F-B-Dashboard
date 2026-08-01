import os
import base64
import requests

BREVO_API_KEY = os.environ["BREVO_API_KEY"]
SENDER_EMAIL = os.environ["SENDER_EMAIL"]  # the address you verified in Brevo
RECIPIENTS = [r.strip() for r in os.environ["RECIPIENTS"].split(",")]

DASHBOARD_URL = "https://suwicha-cst.github.io/F-B-Dashboard/"

with open("dashboard.png", "rb") as f:
    image_b64 = base64.b64encode(f.read()).decode("utf-8")

html = f"""
<html>
  <body style="font-family: Arial, sans-serif;">
    <p>Good morning,</p>
    <p>Here's today's snapshot of the Jul's &amp; Zephyr dashboard:<br>
       For the live, interactive version where you can filter by outlet, date, or period:
       <a href="{DASHBOARD_URL}">click here</a></p>
    <img src="cid:dashboard.png" style="max-width:700px; width:100%; border:1px solid #ddd;" />
  </body>
</html>
"""

payload = {
    "sender": {"email": SENDER_EMAIL},
    "to": [{"email": r} for r in RECIPIENTS],
    "subject": "Jul's & Zephyr — Daily Performance Snapshot",
    "htmlContent": html,
    "attachment": [
        {"content": image_b64, "name": "dashboard.png"}
    ],
}

response = requests.post(
    "https://api.brevo.com/v3/smtp/email",
    headers={
        "api-key": BREVO_API_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json",
    },
    json=payload,
)

if response.status_code >= 300:
    raise Exception(f"Brevo API error {response.status_code}: {response.text}")

print("Email sent successfully via Brevo to:", RECIPIENTS)
print("Response:", response.json())
