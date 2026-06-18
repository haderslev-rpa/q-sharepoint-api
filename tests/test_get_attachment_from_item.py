# tests/test_download_item_attachments.py

"""
TEST: Hent attachments fra SharePoint liste-item
"""

import os

from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_attachments import (
    list_attachments_for_item,
    download_attachment
)


# -------------------------------------------------
# INIT (miljø – meget vigtigt)
# -------------------------------------------------
AutomationServer.from_environment()


# -------------------------------------------------
# KONFIG (ret kun disse værdier)
# -------------------------------------------------

SITE_NAME = "Automatisering"      # SharePoint site-navn
LIST_TITLE = "Test - Rune"           # Liste-navn
ITEM_ID = 7                      # Item-id (tal)

TARGET_FOLDER = "/home/dirujo/attachments_test"


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_download_attachments():
    print("🚀 Starter test: download attachments\n")

    # 1) Hent liste over attachments
    attachments = list_attachments_for_item(
        site_name=SITE_NAME,
        list_title=LIST_TITLE,
        item_id=ITEM_ID
    )

    print(f"✅ Fundet {len(attachments)} attachments")

    if not attachments:
        print("⚠️ Ingen attachments fundet – testen stopper")
        return

    # 2) Download hvert attachment
    for attachment in attachments:
        local_path = download_attachment(
            site_name=SITE_NAME,
            attachment=attachment,
            target_folder=TARGET_FOLDER
        )

        print("⬇️ Gemt lokalt:", local_path)

        # 3) Valider at filen findes
        assert os.path.exists(local_path), (
            f"❌ Filen blev ikke gemt: {local_path}"
        )

    print("\n🎉 Test gennemført – alle filer hentet korrekt")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_download_attachments()
#og gem dem lokalt på Linux