# q_sharepoint_api/sp_list_schema.py

from q_sharepoint_api.sp_api import get_client  # funktion (hent klient)


# -------------------------------------------------
# FELTER SOM ALDRIG MÅ BRUGES (UI-navne)
# -------------------------------------------------

IGNORED_UI_FIELDS = {
    "Antal underordnede elementer",
    "Antal underordnede mapper",
    "App oprettet af",
    "App ændret af",
    "Element er en post",
    "Farvemærke",
    "Id",                    # stort I (systemfelt)
    "Indholdstype",
    "Indstilling for mærkat",
    "Mærkat anvendt af",
    "Opbevaringsmærkat",
    "Opbevaringsmærkat er anvendt",
    "Overholdelsesaktiv-id",
    "Rediger",
    "Titel",                 # dansk titel (må ikke bruges)
    "Type",
    "Version",
    "Vedhæftede filer",
    "Ændret",
    "Ændret af",
    "Oprettet",
    "Oprettet af",
}


def get_list_schema(site_name, list_name):
    """
    Returnerer schema for SharePoint-liste
    Kun tilladte forretningsfelter
    """

    client = get_client()                         # objekt (SharePoint-klient)
    site_id = client.get_site_id(site_name)       # tekst (site-id)
    list_id = client.get_list_id(site_id, list_name)

    schema = {}                                   # dict (schema)

    for col in client.get_list_columns(site_id, list_id):

        ui_name = col.get("displayName")          # tekst (UI-navn)

        # spring ignorerede felter over
        if ui_name in IGNORED_UI_FIELDS:
            continue

        schema[ui_name] = {
            "api_name": col.get("name"),          # tekst (API-navn)
            "type": col.get("columnType"),        # tekst (datatype)
            "read_only": col.get("readOnly", False)
        }

    return schema