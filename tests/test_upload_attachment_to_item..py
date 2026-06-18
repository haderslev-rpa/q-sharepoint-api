# tests/test_upload_attachment_to_item.py

"""
TEST:
Download attachment fra et SharePoint liste-item
og upload SAMME fil som attachment til et item
"""

import os

from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_attachments import (
    list_attachments_for_item,
    download_attachment
)

from q_sharepoint_api.functionality.sp_list_attachments import add_attachment_to_item


# -------------------------------------------------
# INIT (miljø – meget vigtigt)
# -------------------------------------------------
AutomationServer.from_environment()


# -------------------------------------------------
# KONFIG (ret kun disse værdier)
# -------------------------------------------------

SITE_NAME = "Automatisering"
LIST_TITLE = "Test - Rune"

SOURCE_ITEM_ID = 7      # item hvor filen hentes fra
TARGET_ITEM_ID = 8      # item hvor filen uploades til

TARGET_FOLDER = "/home/dirujo/attachments_test"


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_download_and_upload_attachment():
    print("🚀 Starter test: download → upload attachment\n")

    # 1) Hent attachments fra source-item
    attachments = list_attachments_for_item(
        site_name=SITE_NAME,
        list_title=LIST_TITLE,
        item_id=SOURCE_ITEM_ID
    )

    print(f"✅ Fundet {len(attachments)} attachments")

    if not attachments:
        print("❌ Ingen attachments at teste med")
        return

    # 2) Download FØRSTE attachment
    attachment = attachments[0]

    local_path = download_attachment(
        site_name=SITE_NAME,
        attachment=attachment,
        target_folder=TARGET_FOLDER
    )

    print("⬇️ Gemt lokalt:", local_path)

    # Valider at filen findes
    assert os.path.exists(local_path), "❌ Filen blev ikke gemt lokalt"

    # 3) Upload SAMME fil til target-item
    print("⬆️ Uploader samme fil som attachment...")

    result = add_attachment_to_item(
        site_name=SITE_NAME,
        list_title=LIST_TITLE,
        item_id=TARGET_ITEM_ID,
        file_path=local_path
    )

    print("✅ Attachment uploadet")
    print("📎 SharePoint svar:", result)

    print("\n🎉 Test gennemført: download + upload OK")


# -------------------------------------------------
# KØR TEST DIREKTE
# -------------------------------------------------

if __name__ == "__main__":
    test_download_and_upload_attachment()