"""
TEST: Hent attachments fra et SharePoint-listeelement
og gem dem lokalt på Linux.
"""

from pathlib import Path

from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_list_attachments import (
    download_attachment,
    list_attachments_for_item,
)


# -------------------------------------------------
# INIT
# -------------------------------------------------

AutomationServer.from_environment()


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

SITE_NAME = "Automatisering"
LIST_NAME = "Test - Rune"
ITEM_ID = 16

TARGET_FOLDER = Path(
    "/home/dirujo/attachments_test"
)


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_download_attachments() -> None:
    """
    Henter alle attachments fra ét SharePoint-listeelement.

    Hver attachment gemmes i TARGET_FOLDER på Linux.
    Funktionen returnerer None, når testen er gennemført.
    """
    print("")
    print("=" * 70)
    print("TESTER DOWNLOAD AF SHAREPOINT-ATTACHMENTS")
    print("=" * 70)

    print(f"Site: {SITE_NAME}")
    print(f"Liste: {LIST_NAME}")
    print(f"Item-id: {ITEM_ID}")
    print(f"Linux-mappe: {TARGET_FOLDER}")
    print("")

    # Opret målmappen på Linux, hvis den ikke allerede findes.
    TARGET_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Funktionen forventer parameteren list_name.
    attachments = list_attachments_for_item(
        site_name=SITE_NAME,
        list_name=LIST_NAME,
        item_id=ITEM_ID,
    )

    print(
        f"Fundet {len(attachments)} attachment(s)."
    )

    if not attachments:
        raise AssertionError(
            f"SharePoint-item {ITEM_ID} i listen "
            f"{LIST_NAME!r} har ingen attachments."
        )

    downloaded_files = []

    for attachment_number, attachment in enumerate(
        attachments,
        start=1,
    ):
        file_name = attachment.get("FileName")

        if not file_name:
            raise AssertionError(
                "SharePoint returnerede en attachment "
                "uden feltet 'FileName'."
            )

        print("")
        print(
            f"{attachment_number}. Downloader: {file_name}"
        )

        local_path = download_attachment(
            site_name=SITE_NAME,
            attachment=attachment,
            target_folder=str(TARGET_FOLDER),
        )

        downloaded_file = Path(local_path)

        assert downloaded_file.is_file(), (
            f"Filen blev ikke gemt: {downloaded_file}"
        )

        assert downloaded_file.stat().st_size > 0, (
            f"Den downloadede fil er tom: {downloaded_file}"
        )

        downloaded_files.append(downloaded_file)

        print(f"   Gemt: {downloaded_file}")
        print(
            f"   Størrelse: "
            f"{downloaded_file.stat().st_size} bytes"
        )

    print("")
    print("=" * 70)
    print("TEST GENNEMFØRT")
    print("=" * 70)

    print(
        f"Downloadet {len(downloaded_files)} fil(er) til:"
    )
    print(TARGET_FOLDER)

    for downloaded_file in downloaded_files:
        print(f" - {downloaded_file.name}")


# -------------------------------------------------
# DIREKTE KØRSEL
# -------------------------------------------------

if __name__ == "__main__":
    test_download_attachments()