import os
import base64
import requests

MAILJET_API_KEY = os.environ["MAILJET_API_KEY"]
MAILJET_SECRET_KEY = os.environ["MAILJET_SECRET_KEY"]
SENDER_EMAIL = os.environ["SENDER_EMAIL"]  # the address you verified in Mailjet
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
    "Messages": [
        {
            "From": {"Email": SENDER_EMAIL, "Name": "Dashboard"},
            "To": [{"Email": r} for r in RECIPIENTS],
            "Subject": "Jul's & Zephyr — Daily Performance Snapshot",
            "HTMLPart": html,
            "InlinedAttachments": [
                {
                    "ContentType": "image/png",
                    "Filename": "dashboard.png",
                    "ContentID": "dashboard_image",
                    "Base64Content": image_b64,
                }
            ],
        }
    ]
}

response = requests.post(
    "https://api.mailjet.com/v3.1/send",
    auth=(MAILJET_API_KEY, MAILJET_SECRET_KEY),
    json=payload,
)

if response.status_code >= 300:
    raise Exception(f"Mailjet API error {response.status_code}: {response.text}")

print("Email sent successfully via Mailjet to:", RECIPIENTS)
print("Response:", response.json())
