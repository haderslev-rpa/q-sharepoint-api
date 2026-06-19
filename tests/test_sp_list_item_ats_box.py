from pprint import pprint
from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_item_ats_box import (
    create_or_update_from_ats_box,
    create_from_ats_box,
    update_from_ats_box
)

# -------------------------------------------------
# INIT
# -------------------------------------------------
AutomationServer.from_environment()

SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"


# -------------------------------------------------
# MOCK ITEM
# -------------------------------------------------
class MockItem:
    def __init__(self, data):
        self.data = data

    def update(self, data):
        self.data = data


# -------------------------------------------------
# DINE TEST DATA
# -------------------------------------------------
def build_test_data():
    return {
        "box": {
            "sharepoint": {
                "id": "",
                "Dato": "2025-12-18T23:00:00Z",
                "Kaffe": True,
                "Links": "",
                "Title": "rwar",
                "Beløb": 200,
                "created": "2025-12-19T09:06:24Z",   # ✅ skal ignoreres
                "Hardware": "Valg 2",
                "modified": "2026-06-10T12:15:14Z",  # ✅ skal ignoreres
                "Godkendt?": True,
                "Robot kommentar": "Test kommentar",
                "Sagsbehandler - Initialer": "test"
            }
        },
        "defer": None,
        "state": [],
        "status": {}
    }


# -------------------------------------------------
# TEST 1: UPDATE (ID findes)
# -------------------------------------------------
def test_update():
    print("\n🔄 TEST: UPDATE\n")

    data = build_test_data()
    item = MockItem(data)

    print("📦 INPUT (box.sharepoint):")
    pprint(item.data["box"]["sharepoint"])

    result = create_or_update_from_ats_box(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item=item,
        robot_kommentar="Opdateret via test"
    )

    print("\n✅ RESULTAT:")
    pprint(result)

    print("\n📦 BOX EFTER:")
    pprint(item.data["box"]["sharepoint"])


# -------------------------------------------------
# TEST 2: CREATE (fjern ID)
# -------------------------------------------------
def test_create():
    print("\n🆕 TEST: CREATE\n")

    data = build_test_data()

    # ✅ Fjern ID → tving CREATE
    data["box"]["sharepoint"].pop("id", None)

    item = MockItem(data)

    print("📦 INPUT (box.sharepoint):")
    pprint(item.data["box"]["sharepoint"])

    result = create_or_update_from_ats_box(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item=item,
        robot_kommentar="Oprettet via test"
    )

    print("\n✅ RESULTAT:")
    pprint(result)

    print("\n📦 BOX EFTER:")
    pprint(item.data["box"]["sharepoint"])


# -------------------------------------------------
# TEST 3: DEBUG – FELTER DER SENDES
# -------------------------------------------------
def test_payload_debug():
    print("\n🧪 TEST: DEBUG PAYLOAD\n")

    data = build_test_data()
    item = MockItem(data)

    sp = item.data["box"]["sharepoint"]

    print("\n📦 ORIGINAL:")
    pprint(sp)

    print("\n🧹 FORVENTET (systemfelter fjernes):")
    cleaned = {
        k: v for k, v in sp.items()
        if k not in ["id", "created", "modified"]
    }
    pprint(cleaned)


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":

    print("\n==============================")
    print(" SHAREPOINT ATS BOX TEST ")
    print("==============================")

    test_payload_debug()
    test_update()
    test_create()

    print("\n✅ TEST FÆRDIG")
