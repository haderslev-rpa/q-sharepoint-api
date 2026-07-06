import requests
from datetime import datetime
import os
import dotenv

from q_sharepoint_api.sp_api import get_client  # funktion (henter client)


def test_calendar():
    """
    Tester Outlook kalender via APP (Graph)
    """

    client = get_client()  # object (client instans)
    headers = client.auth.graph_headers()

    # ✅ VIGTIGT – samme som Blue Prism
    headers["Prefer"] = 'outlook.timezone="W. Europe Standard Time"'

    # 🔧 Mailbox (samme som du brugte i Blue Prism)
    dotenv.load_dotenv()
    user_mail = os.getenv("MAIL_AKASSE")

    # 🔧 Dato (samme format som Blue Prism laver)
    start = "2026-01-01T00:00:00"
    end = "2026-05-01T00:00:00"

    # ✅ STEP 1: hent default calendar
    url_cal = f"https://graph.microsoft.com/v1.0/users/{user_mail}/calendars"

    print("🔍 Henter kalendere...\n")

    r = requests.get(url_cal, headers=headers, timeout=30)
    print("STATUS calendars:", r.status_code)
    r.raise_for_status()

    calendars = r.json().get("value", [])

    # find default kalender
    default_cal = next((c for c in calendars if c.get("isDefaultCalendar")), None)

    if not default_cal:
        print("❌ Ingen default kalender fundet")
        return

    calendar_id = default_cal["id"]

    print("✅ Default kalender fundet:", calendar_id)

    # ✅ STEP 2: hent kalender events (samme som BP)
    url_events = (
        f"https://graph.microsoft.com/v1.0/users/{user_mail}"
        f"/calendars/{calendar_id}"
        f"/calendarView"
        f"?startDateTime={start}"
        f"&endDateTime={end}"
    )

    print("\n🔍 Henter møder...\n")

    r = requests.get(url_events, headers=headers, timeout=30)

    print("STATUS events:", r.status_code)
    print(r.text[:500])

    r.raise_for_status()

    events = r.json().get("value", [])

    print("\n✅ RESULTAT:\n")

    for e in events:
        subject = e.get("subject")
        start_dt = e.get("start", {}).get("dateTime")
        tz = e.get("start", {}).get("timeZone")

        print(f"📅 {subject}")
        print(f"   Start: {start_dt}")
        print(f"   TimeZone: {tz}")
        print("-" * 40)


if __name__ == "__main__":
    test_calendar()