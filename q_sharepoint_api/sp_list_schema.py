# q_sharepoint_api/sp_list_schema.py

from q_sharepoint_api.sp_api import get_client


IGNORED_UI_FIELDS = {
    "Antal underordnede elementer",
    "Antal underordnede mapper",
    "App oprettet af",
    "App ændret af",
    "Element er en post",
    "Farvemærke",
    "Id",
    "Indholdstype",
    "Indstilling for mærkat",
    "Mærkat anvendt af",
    "Opbevaringsmærkat",
    "Opbevaringsmærkat er anvendt",
    "Overholdelsesaktiv-id",
    "Rediger",
    "Titel",
    "Type",
    "Version",
    "Vedhæftede filer",
    "Ændret",
    "Ændret af",
    "Oprettet",
    "Oprettet af",
}


def get_list_schema(site_name, list_name):

    client = get_client()
    site_id = client.get_site_id(site_name)
    list_id = client.get_list_id(site_id, list_name)

    schema = {}

    for col in client.get_list_columns(site_id, list_id):

        ui_name = col.get("displayName")

        if ui_name in IGNORED_UI_FIELDS:
            continue

        # ✅ SIMPEL TYPE (kun det der virker stabilt)
        col_type = "text"

        if col.get("choice"):
            col_type = "choice"
        elif col.get("personOrGroup"):
            col_type = "user"
        elif col.get("lookup"):
            col_type = "lookup"
        elif col.get("dateTime"):
            col_type = "date"

        schema[ui_name] = {
            "api_name": col.get("name"),
            "type": col_type,
            "read_only": col.get("readOnly", False)
        }

    return schema