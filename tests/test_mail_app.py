import os
import requests
from dotenv import load_dotenv

from q_sharepoint_api.sp_api import get_client  # funktion (henter client)


def test_mail_app():
    """
    Tester mail via APP login (Graph)
    """

    # Henter variabler fra .env filen
    load_dotenv()

    client = get_client()  # objekt (konkret instans)

    headers = client.auth.graph_headers()

    user_mail = os.getenv("TEST_MAIL")

    if not user_mail:
        raise Exception("TEST_MAIL mangler i .env filen")

    url = f"https://graph.microsoft.com/v1.0/users/{user_mail}/messages?$top=3"

    print("🔍 Tester mail via APP...\n")
    print("Tester postkasse:", user_mail)

    response = requests.get(url, headers=headers, timeout=30)

    print("STATUS:", response.status_code)
    print(response.text[:300])

    response.raise_for_status()

    data = response.json()

    print("\n✅ SUCCESS – mails:\n")

    for m in data.get("value", []):
        subject = m.get("subject")
        sender = m.get("from", {}).get("emailAddress", {}).get("address")

        print(f"📧 Fra: {sender}")
        print(f"   Emne: {subject}")
        print("-" * 40)


if __name__ == "__main__":
    test_mail_app()