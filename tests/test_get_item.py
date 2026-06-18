# tests/test_get_item_json.py

from automation_server_client import AutomationServer
from pprint import pprint
from q_sharepoint_api.functionality.sp_listitem_json import get_item_as_json

# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------
AutomationServer.from_environment()


# -------------------------------------------------
# KONFIG
# -------------------------------------------------
SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"
ITEM_ID = 7
include_raw_json = True    # inkluder raw JSON i output?


# -------------------------------------------------
# TEST
# -------------------------------------------------
def test_get_item_json():
    print("🚀 Henter item som rå JSON\n")

    data = get_item_as_json(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item_id=ITEM_ID,
        include_raw=include_raw_json
    )

    print("✅ Rå JSON fra SharePoint:\n")
    pprint(data)   # pæn print af JSON


if __name__ == "__main__":
    test_get_item_json()
