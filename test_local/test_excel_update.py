# tests/test_graph_excel_update.py

"""
TEST: REDIGER EXCEL DIREKTE I SHAREPOINT (GRAPH API)

Denne test demonstrerer hvordan du:
- finder en Excel fil i SharePoint
- opretter en Excel session (vigtigt!)
- skriver data direkte i et worksheet

🚨 VIGTIGT:
Denne test skriver DIREKTE i Excel filen!
Brug en test-fil – ikke produktionsdata!

-------------------------------------------------

FLOW (arbejdsgang):
1. Opret SharePoint klient (get_client)
2. Find site_id
3. Find file_id (Excel fil)
4. Opret Excel session
5. Opdater celler i Excel

-------------------------------------------------

KRAV:
- Filen SKAL ligge i SharePoint dokumentbibliotek
- Du skal kende:
    - site navn
    - fil sti (ikke UNC!)
    - sheet navn

-------------------------------------------------

EKSEMPEL FILE PATH:
"Shared Documents/Test/test.xlsx"

IKKE:
"\\server\mappe\fil.xlsx"
"""

from automation_server_client import AutomationServer  # modul (ATS init)
from q_sharepoint_api.sp_api import get_client         # funktion (hent klient)


# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------
AutomationServer.from_environment()  # funktion (init ATS)


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

SITE_NAME = "Automatisering"   # variabel (SharePoint site navn)

FILE_PATH = "Shared Documents/Test/test.xlsx"  
# variabel (sti i SharePoint - IKKE Windows path)

SHEET_NAME = "Ark1"  # variabel (Excel ark navn)

CELL_RANGE = "A1:C1"  # variabel (celle område)


# DATA som skal skrives
VALUES = [
    ["Rune", "Haderslev", 999]
]
# liste (2D liste til Excel)


# -------------------------------------------------
# TEST
# -------------------------------------------------

def test_update_excel():
    print("🚀 Starter Excel Graph test\n")

    # 1. Opret klient
    client = get_client()  # objekt (SharePointClient)

    print("✅ Klient oprettet")

    # 2. Find site_id
    site_id = client.get_site_id(SITE_NAME)  # funktion (hent site id)

    print(f"✅ Site fundet: {SITE_NAME}")
    print(f"   site_id: {site_id}\n")

    # 3. Find file_id
    file_id = client.get_drive_item_id(site_id, FILE_PATH)  # funktion (hent fil id)

    print("✅ Excel fil fundet")
    print(f"   file_id: {file_id}\n")

    # 4. Opret Excel session
    session_id = client.create_excel_session(file_id)  # funktion (opret session)

    print("✅ Excel session oprettet")
    print(f"   session_id: {session_id}\n")

    # 5. Opdater Excel
    result = client.update_excel_range(
        file_id=file_id,
        sheet_name=SHEET_NAME,
        cell_range=CELL_RANGE,
        values=VALUES,
        session_id=session_id
    )

    print("✅ Excel opdateret!\n")

    print("📄 Resultat fra Graph API:")
    print(result)


# -------------------------------------------------
# RUN TEST MANUELT
# -------------------------------------------------

if __name__ == "__main__":
    test_update_excel()
