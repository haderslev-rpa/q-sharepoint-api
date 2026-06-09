# q_sharepoint_api/sp_list_items.py

"""
HENT FLERE SHAREPOINT LIST ITEMS (LAV-NIVEAU)

Denne fil bruges når:
- en proces starter i SharePoint
- robotten skal finde nye / eksisterende items
- man vil hente flere items ad gangen

⚠️ VIGTIGT:
- Denne fil returnerer ALDRIG Automation Server item
- Den returnerer KUN SharePoint-data
- Processer bør normalt IKKE bruge denne direkte
  men i stedet via ATS-wrapper (hent_..._til_box)

Bruges typisk:
- i start af en proces - Feed Queue
- i scripts / tests
"""

from q_sharepoint_api.sp_api import get_client          # funktion (hent SP-klient)
from q_sharepoint_api.sp_list_schema import get_list_schema  # funktion (schema-filter)


def get_list_items(site_name, list_name, item_id=None, include_raw=False):
    """
    Henter SharePoint list items som REN forretnings-JSON.

    Parametre:
    - site_name (tekst): SharePoint site
    - list_name (tekst): SharePoint liste
    - item_id (valgfri): hvis sat → hent kun dette item
    - include_raw (bool): hvis True → medtag rå Graph-svar

    Returnerer:
    {
        "exists": bool,
        "count": int,
        "items": [ { sharepoint_item }, ... ]
    }
    """

    # Opret SharePoint-klient (objekt – API-adgang)
    client = get_client()

    # Find site-id (tekst – Graph-id)
    site_id = client.get_site_id(site_name)

    # Find list-id ud fra listens UI-navn
    list_id = client.get_list_id(site_id, list_name)

    # Hent schema (kun tilladte forretningsfelter)
    schema = get_list_schema(site_name, list_name)

    # Hent ALLE items råt fra SharePoint (Graph)
    items_raw = client.get_list_items_raw(site_id, list_id)

    items = []  # liste (samling af items)

    for item in items_raw:

        # Hvis vi kun vil have ét bestemt item
        if item_id and item["id"] != str(item_id):
            continue

        fields = item["fields"]  # dict (kolonne-værdier)

        # Basis-felter som ALTID er med
        ui_item = {
            "id": item["id"],  # SharePoint item-id (vigtigt!)
            "created": item.get("createdDateTime"),
            "modified": item.get("lastModifiedDateTime")
        }

        # Tilføj kun felter fra schema (ingen system-støj)
        for ui_name, meta in schema.items():
            api_name = meta["api_name"]
            ui_item[ui_name] = fields.get(api_name, "")

        items.append(ui_item)

    result = {
        "exists": len(items) > 0,
        "count": len(items),
        "items": items
    }

    # Bruges KUN til debug / test
    if include_raw:
        result["raw"] = items_raw

    return result
