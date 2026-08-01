import os
import base64
import requests

RESEND_API_KEY = os.environ["RESEND_API_KEY"]
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
    <img src="cid:dashboard_image" style="max-width:700px; width:100%; border:1px solid #ddd;" />
  </body>
</html>
"""

payload = {
    "from": "Dashboard <onboarding@resend.dev>",
    "to": RECIPIENTS,
    "subject": "Jul's & Zephyr — Daily Performance Snapshot",
    "html": html,
    "attachments": [
        {
            "filename": "dashboard.png",
            "content": image_b64,
            "content_id": "dashboard_image",
        }
    ],
}

response = requests.post(
    "https://api.resend.com/emails",
    headers={
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json",
    },
    json=payload,
)

if response.status_code >= 300:
    raise Exception(f"Resend API error {response.status_code}: {response.text}")

print("Email sent successfully via Resend to:", RECIPIENTS)
print("Response:", response.json())
