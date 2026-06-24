from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.sp_list_schema import get_list_schema
from q_sharepoint_api.functionality.sp_list_items import get_list_items
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

    # -------------------------------------------------
    # ✅ VALIDERING FØR LOOP 
    # -------------------------------------------------
    invalid_fields = [
        k for k in data.keys()
        if k not in schema and not k.startswith("_") and k != "id"
    ]

    if invalid_fields:
        raise Exception(
            f"Ugyldige felter i input: {invalid_fields}\n"
            f"Gyldige felter i SharePoint: {list(schema.keys())}"
        )

    payload = {}

    existing = None
    if item_id:
        existing = get_list_items(site_name, list_name, item_id)["items"][0]

    for ui_name, value in data.items():

        # ✅ INTERN FELT - spring over
        if ui_name == "_robot_comment":
            continue

        # -------------------------------------------------
        # ✅ DATO (robust – køres FØR alt andet)
        # -------------------------------------------------
        if isinstance(value, str) and "T" in value:
            try:
                value = value.split("T")[0]
            except:
                pass

        if ui_name == "id":
            continue

        if ui_name not in schema:
            raise Exception(f"Felt '{ui_name}' findes ikke i schema")

        meta = schema[ui_name]

        if meta.get("type") == "hyperlink":
            raise Exception(
                f"Felt '{ui_name}' er en SharePoint hyperlink kolonne. "
                f"Disse understøttes ikke ved create/update via API. "
                f"Fjern feltet fra input."
            )

        if meta.get("read_only"):
            continue

        api_name = meta["api_name"]

        # -------------------------------------------------
        # ✅ DATO (dateTime kolonne)
        # -------------------------------------------------
        # ✅ ROBUST check – virker uanset schema variation
        if isinstance(value, str) and "T" in value:
            payload[api_name] = value.split("T")[0]
            continue

        # -------------------------------------------------
        # ✅ Robot kommentar (append)
        # -------------------------------------------------
        if ui_name == ROBOT_COMMENT_UI_NAME:

            incoming = data.get("_robot_comment")

            if incoming:
                ts = dk_timestamp()
                old = existing.get(ROBOT_COMMENT_UI_NAME, "") if existing else ""

                if old:
                    payload[api_name] = f"{old}\n{ts}: {incoming}"
                else:
                    payload[api_name] = f"{ts}: {incoming}"

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
        # ✅ ignorer tomme værdier (men IKKE ved CREATE)
        # -------------------------------------------------
        if value in ("", None):
            # Ved UPDATE → ok at skippe
            if not item_id:
                # CREATE: behold tom værdi hvis field findes i schema
                payload[api_name] = value if value is not None else ""
            continue

        # -------------------------------------------------
        # ✅ almindelig værdi
        # -------------------------------------------------
        payload[api_name] = value

    # ✅ Her bruges det RIGTIGE item_id (fra starten)
    try:
        if item_id:
            client.update_list_item(site_id, list_id, item_id, payload)
        else:
            created = client.create_list_item(
                site_id,
                list_id,
                {"fields": payload}
            )
            item_id = created["id"]

    except Exception as e:

        from pprint import pprint

        print("\n" + "=" * 100)
        print("❌ SHAREPOINT ERROR DEBUG")
        print("=" * 100)

        # -------------------------------------------------
        # 🔹 INPUT
        # -------------------------------------------------
        print("\n🔹 INPUT (box.sharepoint):")
        pprint(data)

        # -------------------------------------------------
        # 🔹 PAYLOAD (det der sendes)
        # -------------------------------------------------
        print("\n🔹 PAYLOAD (sendt til SharePoint):")
        pprint(payload)

        # -------------------------------------------------
        # 🔹 SCHEMA
        # -------------------------------------------------
        print("\n🔹 SCHEMA (tilladte felter):")
        for k, v in schema.items():
            print(f" - {k:30} → api: {v['api_name']} | type: {v['type']}")

        # -------------------------------------------------
        # 🔹 FELTER DER IKKE FINDES I SCHEMA
        # -------------------------------------------------
        print("\n🔹 FELTER DER IKKE MATCHER SCHEMA:")

        unknown_fields = [
            k for k in data.keys()
            if k not in schema and not k.startswith("_")
        ]

        if unknown_fields:
            for k in unknown_fields:
                print(f" - {k}")
        else:
            print(" ✅ ingen")

        # -------------------------------------------------
        # 🔹 FELTER DU MANGLER (BONUS)
        # -------------------------------------------------
        print("\n🔹 FELTER DU IKKE HAR SENDT (kan være required):")
        missing_fields = [k for k in schema.keys() if k not in data]
        if missing_fields:
            for k in missing_fields:
                print(f" - {k}")
        else:
            print(" ✅ ingen")

        # -------------------------------------------------
        # 🔹 FELTER MED None / tomme (ofte årsag)
        # -------------------------------------------------
        print("\n🔹 FELTER MED TOMME VÆRDIER:")
        empty_fields = [
            k for k, v in data.items()
            if v in ("", None) and not k.startswith("_")
        ]

        if empty_fields:
            for k in empty_fields:
                print(f" - {k}")
        else:
            print(" ✅ ingen")

        # -------------------------------------------------
        # 🔹 PAYLOAD vs DATA mismatch
        # -------------------------------------------------
        print("\n🔹 FELTER DER FORSVINDER FRA DATA → PAYLOAD:")
        for k in data.keys():
            if k in schema:
                api_name = schema[k]["api_name"]
                if api_name not in payload:
                    print(f" - {k}")

        # -------------------------------------------------
        # 🔹 ORIGINAL ERROR
        # -------------------------------------------------
        print("\n🔹 ORIGINAL ERROR:")
        print(str(e))

        print("=" * 100 + "\n")

        raise

    return get_list_items(site_name, list_name, item_id)
