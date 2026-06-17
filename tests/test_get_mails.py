import requests

from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.copilot_auth import get_token


# -------------------------------------------------
# SETUP
# -------------------------------------------------
token = get_token()

headers = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "application/json"
}


# -------------------------------------------------
# HENT MAILS
# -------------------------------------------------
mail = "dirxhel@haderslev.dk"

url = f"https://graph.microsoft.com/v1.0/users/{mail}/messages?$top=5"

print("🔍 Henter mails...\n")

response = requests.get(url, headers=headers)
response.raise_for_status()

data = response.json()


# -------------------------------------------------
# VIS RESULTAT
# -------------------------------------------------
messages = data.get("value", [])

print("✅ RESULTAT:\n")

for m in messages:
    subject = m.get("subject")
    sender = m.get("from", {}).get("emailAddress", {}).get("address")

    print(f"📧 Fra: {sender}")
    print(f"   Emne: {subject}")
    print("-" * 40)