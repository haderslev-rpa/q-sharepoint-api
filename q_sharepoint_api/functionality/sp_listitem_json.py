# q_sharepoint_api/sp_listitem_json.py

from q_sharepoint_api.sp_api import get_client


def get_item_as_json(site_name, list_name, item_id, include_raw=False):
    """
    Henter et SharePoint liste-item som JSON.
    Hvis include_raw=True returneres også raw Graph-data.
    """

    client = get_client()                    # client (Graph-klient)
    site_id = client.get_site_id(site_name)  # site-id
    list_id = client.get_list_id(site_id, list_name)  # ✅ STANDARD

    items = client.get_items(site_id, list_id)

    for item in items:
        if item["id"] == str(item_id):
            parsed = item["fields"]   # parsed (kolonne-data)
            raw = item                # raw (fuldt Graph-svar)

            if include_raw:
                return {
                    "parsed": parsed,
                    "raw": raw
                }

            return parsed

    raise Exception(f"Item {item_id} ikke fundet")