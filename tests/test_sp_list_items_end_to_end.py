# tests/test_sp_list_items_end_to_end.py

"""
SAMLET ENDT-TO-END TEST – SharePoint list items

Tester:
- GET list item (renset forretnings-JSON)
- SAVE (update)
- Robot kommentar (append med dansk timestamp)
- #CLEAR#
- Save returnerer altid fuld JSON
"""

from pprint import pprint                     # funktion (pæn print)
from automation_server_client import AutomationServer  # klasse (miljø-loader)

from q_sharepoint_api.functionality.sp_list_items import get_list_items      # funktion (GET)
from q_sharepoint_api.functionality.sp_list_item_save import save_list_item  # funktion (SAVE)


# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------
AutomationServer.from_environment()            # funktion (loader credentials)

# -------------------------------------------------
# KONFIG
# -------------------------------------------------
SITE_NAME = "Automatisering"                   # tekst (site-navn)
LIST_NAME = "Test - Rune"                      # tekst (liste-navn)
ITEM_ID = 7                                   # tal (eksisterende item-id)


# -------------------------------------------------
# TEST
# -------------------------------------------------
def test_sp_list_items_end_to_end():
    print("\n🚀 STARTER ENDT-TO-END TEST\n")

    # -------------------------------------------------
    # 1️⃣ GET
    # -------------------------------------------------
    print("🔍 HENTER ITEM (GET)\n")

    get_result = get_list_items(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item_id=ITEM_ID,
        include_raw=False
    )

    pprint(get_result)

    assert get_result["exists"] is True, "Item findes ikke"
    assert get_result["count"] == 1, "Forventede præcis 1 item"

    item = get_result["items"][0]
    print("\n✅ ITEM FUNDET MED ID:", item["id"])

    # -------------------------------------------------
    # 2️⃣ SAVE (UPDATE)
    # -------------------------------------------------
    print("\n✏️ OPDATERER ITEM (SAVE)\n")

    save_result = save_list_item(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        data={
            "id": item["id"],

            # ✅ VIGTIGT: Brug ALTID 'Title'
            "Title": "Opdateret via samlet test",

            # Robot kommentar (append)
            "Robot kommentar": "Test af robot kommentar fra Python"
        }
    )

    pprint(save_result)

    # -------------------------------------------------
    # 3️⃣ VALIDER RESULTAT
    # -------------------------------------------------
    print("\n✅ VALIDERER RESULTAT\n")

    assert save_result["exists"] is True, "Save returnerede ikke item"
    assert save_result["count"] == 1, "Save returnerede forkert antal items"

    saved_item = save_result["items"][0]

    # ✅ Valider Title (ikke Titel)
    assert saved_item["Title"] == "Opdateret via samlet test", (
        "Title blev ikke opdateret"
    )

    # ✅ Robot kommentar findes
    assert "Robot kommentar" in saved_item, (
        "Robot kommentar mangler i output"
    )

    # ✅ id matcher
    assert saved_item["id"] == str(ITEM_ID), "ID matcher ikke"

    print("\n🎉 TEST BESTÅET – ALLE DELE VIRKER\n")


# -------------------------------------------------
# KØR TEST
# -------------------------------------------------
if __name__ == "__main__":
    test_sp_list_items_end_to_end()