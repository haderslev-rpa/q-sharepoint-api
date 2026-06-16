from automation_server_client import AutomationServer

from q_sharepoint_api.sp_listitem_json import (
    get_item_as_json,
    update_item_from_json
)

# -------------------------------------------------
# INIT
# -------------------------------------------------
AutomationServer.from_environment()

# -------------------------------------------------
# KONFIG
# -------------------------------------------------
SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"
ITEM_ID = 7

# -------------------------------------------------
# TEST
# -------------------------------------------------
def test_update_item():
    print("🚀 Tester update_item_from_json\n")

    before = get_item_as_json(SITE_NAME, LIST_NAME, ITEM_ID)
    print("🔎 Før:", before)

    update_item_from_json(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item_id=ITEM_ID,
        json_data={
            "Title": "Opdateret via standard JSON",
            "Beløb": 250,
            "Status": "Godkendt",
            "StartDato": "2026-06-09T00:00:00Z",
            "Kunde": {"lookup_id": 5},
            "Ansvarlig": {"user_id": 23}
        }
    )

    after = get_item_as_json(SITE_NAME, LIST_NAME, ITEM_ID)
    print("✅ Efter:", after)


if __name__ == "__main__":
    test_update_item()