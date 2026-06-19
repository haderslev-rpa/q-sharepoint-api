from pprint import pprint
from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_item_ats_box import (
    get_to_ats_box
)

# -------------------------------------------------
# INIT
# -------------------------------------------------
AutomationServer.from_environment()

SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"


class MockItem:
    def __init__(self):
        self.data = {}

    def update(self, data):
        self.data = data


# -------------------------------------------------
# TEST: GET
# -------------------------------------------------
def test_get():

    print("\n📥 TEST GET (kræver id)\n")

    item = MockItem()

    get_to_ats_box(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        list_item_id=8,  # ✅ eksisterende id
        item=item
    )

    print("\n✅ RESULTAT:")
    pprint(item.data["box"]["sharepoint"])


if __name__ == "__main__":
    test_get()
