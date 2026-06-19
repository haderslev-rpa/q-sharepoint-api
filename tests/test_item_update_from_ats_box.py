from pprint import pprint
from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_item_ats_box import (
    create_or_update_from_ats_box
)

# -------------------------------------------------
# INIT
# -------------------------------------------------
AutomationServer.from_environment()

SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"


class MockItem:
    def __init__(self, data):
        self.data = data

    def update(self, data):
        self.data = data


# -------------------------------------------------
# TEST: UPDATE (kræver id)
# -------------------------------------------------
def test_update():

    print("\n🔄 TEST UPDATE (med id)\n")

    data = {
        "box": {
            "sharepoint": {
                "id": "8",  # ✅ SKAL være et eksisterende id
                "Dato": "2025-12-18T23:00:00Z",
                "Kaffe": True,
                #"Links": "",
                "Title": "TEST UPDATE",
                "Beløb": 999,
                "Hardware": "Valg 2",
                "Godkendt?": True,
                "Robot kommentar": "Eksisterende kommentar",
                "Sagsbehandler - Initialer": "test"
            }
        }
    }

    item = MockItem(data)

    print("📦 INPUT:")
    pprint(item.data["box"]["sharepoint"])

    result = create_or_update_from_ats_box(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item=item,
        robot_kommentar="TEST UPDATE"
    )

    print("\n✅ RESULTAT:")
    pprint(result)

    print("\n📦 BOX EFTER:")
    pprint(item.data["box"]["sharepoint"])


if __name__ == "__main__":
    test_update()