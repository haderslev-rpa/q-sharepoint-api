"""
Integrationstest af de generelle SharePoint-dokumentfunktioner.

Placering:
    tests/test_sp_documents.py

Testfil:
    tests/Robot test.pdf

Kør fra projektets rodmappe:
    uv run python tests/test_sp_documents.py
"""

from __future__ import annotations

from pathlib import Path
from pprint import pprint
from uuid import uuid4

from automation_server_client import AutomationServer

from q_sharepoint_api.functionality.sp_documents import (
    delete_document,
    download_document,
    get_documents,
    upload_document,
)


# ------------------------------------------------------------
# KONFIGURATION
# ------------------------------------------------------------
SITE_NAME = "Automatisering"
LIBRARY_NAME = "Digitalpost"
TEST_FILE_NAME = "Robot test.pdf"


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def get_test_file_path() -> Path:
    """
    Finder Robot test.pdf i samme mappe som testfilen.

    Output:
        Path til tests/Robot test.pdf.
    """
    file_path = Path(__file__).resolve().parent / TEST_FILE_NAME

    if not file_path.is_file():
        raise FileNotFoundError(
            f"Testfilen blev ikke fundet: {file_path}"
        )

    return file_path


def validate_test_pdf(pdf_content: bytes) -> None:
    """
    Kontrollerer, at testfilen indeholder PDF-data.

    Output:
        Returnerer None, hvis filen består kontrollen.
    """
    if not pdf_content:
        raise AssertionError("Robot test.pdf er tom.")

    if not pdf_content.startswith(b"%PDF-"):
        raise AssertionError(
            "Robot test.pdf har ikke en gyldig PDF-signatur."
        )


# ------------------------------------------------------------
# TEST
# ------------------------------------------------------------
def test_sp_documents() -> None:
    """
    Tester upload, kolonner, søgning, download og sletning.

    Output:
        Returnerer None, når alle kontroller er bestået.

    Oprydning:
        Den uploadede testfil slettes altid i finally-blokken,
        hvis uploaden nåede at returnere et drive_item_id.
    """
    print("")
    print("=" * 70)
    print("TESTER SHAREPOINT-DOKUMENTFUNKTIONER")
    print("=" * 70)

    # Initialiserer adgang til Automation Server credentials.
    AutomationServer.from_environment()

    test_file_path = get_test_file_path()
    original_content = test_file_path.read_bytes()
    validate_test_pdf(original_content)

    forsendelses_id = str(uuid4())
    dokument_id = str(uuid4())

    # Det fysiske filnavn i SharePoint er dokument_id.pdf.
    sharepoint_file_name = f"{dokument_id}.pdf"

    drive_item_id: str | None = None

    print(f"Testfil: {test_file_path}")
    print(f"forsendelses_id: {forsendelses_id}")
    print(f"dokument_id: {dokument_id}")

    try:
        # ----------------------------------------------------
        # 1. UPLOAD
        # ----------------------------------------------------
        print("\n1. Uploader dokumentet...")

        upload_result = upload_document(
            site_name=SITE_NAME,
            library_name=LIBRARY_NAME,
            file_name=sharepoint_file_name,
            content=original_content,
            fields={
                "dokument_id": dokument_id,
                "forsendelses_id": forsendelses_id,
            },
            conflict_behavior="fail",
        )

        pprint(upload_result)

        drive_item_id = upload_result.get("drive_item_id")

        assert drive_item_id, (
            "upload_document() returnerede ikke drive_item_id."
        )
        assert upload_result.get("list_item_id"), (
            "upload_document() returnerede ikke list_item_id."
        )
        assert upload_result.get("file_name") == sharepoint_file_name, (
            "Det returnerede SharePoint-filnavn er ikke dokument_id.pdf."
        )

        print("OK: Dokumentet blev uploadet.")

        # ----------------------------------------------------
        # 2. SØG VIA FORSENDELSES-ID
        # ----------------------------------------------------
        print("\n2. Søger via forsendelses_id...")

        forsendelse_result = get_documents(
            site_name=SITE_NAME,
            library_name=LIBRARY_NAME,
            filters={
                "forsendelses_id": forsendelses_id,
            },
            page_size=20,
            max_items=10,
        )

        pprint(forsendelse_result)

        assert forsendelse_result["exists"] is True, (
            "Dokumentet blev ikke fundet via forsendelses_id."
        )
        assert forsendelse_result["count"] == 1, (
            "Forventede præcis ét dokument med testens "
            "forsendelses_id."
        )

        found_document = forsendelse_result["documents"][0]

        assert (
            found_document.get("forsendelses_id")
            == forsendelses_id
        )
        assert found_document.get("dokument_id") == dokument_id
        assert found_document.get("drive_item_id") == drive_item_id
        assert (
            found_document.get("file_name")
            == sharepoint_file_name
        )

        print("OK: Dokumentet blev fundet via forsendelses_id.")

        # ----------------------------------------------------
        # 3. SØG VIA DOKUMENT-ID
        # ----------------------------------------------------
        print("\n3. Søger via dokument_id...")

        dokument_result = get_documents(
            site_name=SITE_NAME,
            library_name=LIBRARY_NAME,
            filters={
                "dokument_id": dokument_id,
            },
            page_size=10,
            max_items=2,
        )

        assert dokument_result["count"] == 1, (
            "Forventede præcis ét dokument med testens dokument_id."
        )
        assert (
            dokument_result["documents"][0].get("drive_item_id")
            == drive_item_id
        )

        print("OK: Dokumentet blev fundet via dokument_id.")

        # ----------------------------------------------------
        # 4. DOWNLOAD VIA SHAREPOINTS DRIVE-ITEM-ID
        # ----------------------------------------------------
        print("\n4. Downloader dokumentet til memory...")

        downloaded_content = download_document(
            site_name=SITE_NAME,
            library_name=LIBRARY_NAME,
            drive_item_id=drive_item_id,
        )

        assert isinstance(downloaded_content, bytes), (
            "download_document() skal returnere bytes."
        )
        assert downloaded_content == original_content, (
            "Den downloadede PDF er ikke identisk med originalen."
        )

        print(
            "OK: Dokumentet blev downloadet korrekt. "
            f"Størrelse: {len(downloaded_content)} bytes."
        )

        # ----------------------------------------------------
        # 5. DOWNLOAD VIA VORES DOKUMENT-ID
        # ----------------------------------------------------
        print("\n5. Downloader dokumentet via dokument_id...")

        downloaded_by_dokument_id = download_document(
            site_name=SITE_NAME,
            library_name=LIBRARY_NAME,
            dokument_id=dokument_id,
        )

        assert downloaded_by_dokument_id == original_content, (
            "Download via dokument_id gav ikke de oprindelige "
            "PDF-bytes."
        )

        print("OK: Dokumentet blev downloadet via dokument_id.")

        print("\n" + "=" * 70)
        print("ALLE TESTS ER BESTÅET")
        print("=" * 70)

    finally:
        # ----------------------------------------------------
        # 6. OPRYDNING
        # ----------------------------------------------------
        if drive_item_id:
            print("\n6. Sletter testdokumentet igen...")

            delete_document(
                site_name=SITE_NAME,
                library_name=LIBRARY_NAME,
                drive_item_id=drive_item_id,
            )

            print(
                "OK: Testdokumentet blev slettet fra SharePoint."
            )
        else:
            print(
                "\nIngen oprydning nødvendig, fordi uploaden ikke "
                "returnerede et drive_item_id."
            )


# ------------------------------------------------------------
# DIREKTE KØRSEL
# ------------------------------------------------------------
if __name__ == "__main__":
    test_sp_documents()