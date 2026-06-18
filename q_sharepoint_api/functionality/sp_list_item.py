# q_sharepoint_api/sp_sharepoint_list_item.py

"""
RENT SHAREPOINT-API (ingen Automation Server)

Brug denne fil når du:
- bare vil hente et SharePoint list item
- bare vil oprette/opdatere et SharePoint list item
- skriver tests eller scripts

⚠️ Denne fil kender IKKE:
- Automation Server item
- item.data
- box
"""

from q_sharepoint_api.functionality.sp_list_items import get_list_items   # funktion (GET items)
from q_sharepoint_api.functionality.sp_list_item_save import save_list_item  # funktion (SAVE)


def get_sharepoint_list_item(site_name, list_name, list_item_id):
    """
    Henter ÉT SharePoint list item som ren dict.
    """

    result = get_list_items(
        site_name=site_name,
        list_name=list_name,
        item_id=list_item_id,
        include_raw=False
    )

    if result["count"] != 1:
        raise Exception("Forventede præcis ét SharePoint list item")

    return result["items"][0]   # dict (rent SharePoint-data)


def create_update_sharepoint_list_item(site_name, list_name, sharepoint_data):
    """
    Opretter eller opdaterer SharePoint list item
    og returnerer ren dict (frisk fra SharePoint).
    """

    result = save_list_item(
        site_name=site_name,
        list_name=list_name,
        data=sharepoint_data
    )

    if result["count"] != 1:
        raise Exception("Save returnerede ikke præcis ét SharePoint list item")

    return result["items"][0]