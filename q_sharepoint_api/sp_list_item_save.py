from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.sp_list_schema import get_list_schema
from q_sharepoint_api.sp_list_items import get_list_items
from q_sharepoint_api.sp_time_utils import dk_timestamp


ROBOT_COMMENT_UI_NAME = "Robot kommentar"
CLEAR_MARKER = "#CLEAR#"

# systemfelter der ALDRIG må gemmes
SYSTEM_FIELDS = {
    "id",
    "created",
    "modified",
    "Author",
    "Editor",
    "_UIVersionString",
}


def save_list_item(site_name, list_name, data):
    """
    Opretter eller opdaterer SharePoint item
    """

    # ✅ AFGØR TYPE FØRST (gem original id)
    item_id = data.get("id")  # variabel (gemmer SharePoint id)

    # ✅ Fjern systemfelter (men BRUG IKKE id herfra efterfølgende)
    clean_data = {}

    for key, value in data.items():
        if key in SYSTEM_FIELDS:
            continue
        clean_data[key] = value

    data = clean_data

    # ✅ Eksplicit validering
    if "Titel" in data:
        raise Exception("Ugyldigt felt 'Titel'. Brug altid 'Title' i JSON.")
    if "Id" in data:
        raise Exception("Ugyldigt felt 'Id'. Brug altid 'id' (lille i).")

    client = get_client()
    site_id = client.get_site_id(site_name)
    list_id = client.get_list_id(site_id, list_name)
    schema = get_list_schema(site_name, list_name)

    payload = {}

    existing = None
    if item_id:
        existing = get_list_items(site_name, list_name, item_id)["items"][0]

    for ui_name, value in data.items():

        if ui_name == "id":
            continue

        if ui_name not in schema:
            raise Exception(
                f"Felt '{ui_name}' findes ikke eller er ikke tilladt i SharePoint"
            )

        meta = schema[ui_name]
        if meta.get("read_only"):
            continue

        api_name = meta["api_name"]

        # -------------------------------------------------
        # ✅ DATO (dateTime kolonne)
        # -------------------------------------------------
        # ✅ ROBUST check – virker uanset schema variation
        if isinstance(value, str) and "T" in value:

            try:
                # Konverter fra SharePoint ISO datetime → date
                payload[api_name] = value.split("T")[0]
                continue
            except:
                pass

        # -------------------------------------------------
        # ✅ HYPERLINK (Links kolonne)
        # -------------------------------------------------
        if meta.get("type") == "hyperlinkOrPicture":

            # Hvis vi får dict fra SharePoint (GET-format)
            if isinstance(value, dict):
                url = value.get("Url")
                desc = value.get("Description", "")

                if url:
                    # SharePoint forventer: "url, beskrivelse"
                    payload[api_name] = f"{url}, {desc}" if desc else url

            # Hvis det allerede er string → brug direkte
            elif isinstance(value, str):
                payload[api_name] = value

            continue

        # -------------------------------------------------
        # ✅ Robot kommentar (append)
        # -------------------------------------------------
        if ui_name == ROBOT_COMMENT_UI_NAME:
            if value == CLEAR_MARKER:
                payload[api_name] = ""
            elif value:
                ts = dk_timestamp()
                old = existing.get(ROBOT_COMMENT_UI_NAME, "") if existing else ""
                payload[api_name] = f"{old}\n{ts}: {value}".strip()
            continue

        # -------------------------------------------------
        # ✅ CLEAR (#CLEAR#)
        # -------------------------------------------------
        if value == CLEAR_MARKER:
            # CREATE → spring over | UPDATE → None
            if item_id:
                payload[api_name] = None
            continue

        # -------------------------------------------------
        # ✅ ignorer tomme værdier
        # -------------------------------------------------
        if value in ("", None):
            continue

        # -------------------------------------------------
        # ✅ almindelig værdi
        # -------------------------------------------------
        payload[api_name] = value

    # ✅ Her bruges det RIGTIGE item_id (fra starten)
    if item_id:
        client.update_list_item(site_id, list_id, item_id, payload)
    else:
        created = client.create_list_item(
            site_id,
            list_id,
            {"fields": payload}
        )
        item_id = created["id"]

    return get_list_items(site_name, list_name, item_id)