import json

from automation_server_client import AutomationServer, Credential

from q_sharepoint_api.sp_auth import SharePointAuth
from q_sharepoint_api.sp_client import SharePointClient
from q_sharepoint_api.sp_generator import generate_mapping
from q_sharepoint_api.sp_mapping import validate_and_map


# -------------------------------------------------
# INIT
# -------------------------------------------------
AutomationServer.from_environment()
cred = Credential.get_credential("API_SHAREPOINT")

cfg = cred.data


# -------------------------------------------------
# INPUT
# -------------------------------------------------
SITE_NAME = "Automatisering"
LIST_NAME = "Rune - Test"


# -------------------------------------------------
# AUTH + CLIENT
# -------------------------------------------------
auth = SharePointAuth(
    tenant_id=cfg["tenant_id"],
    client_id=cfg["client_id"],
    client_secret=cred.password,
    scope=cfg.get("scope", "https://graph.microsoft.com/.default")
)

client = SharePointClient(auth, cfg["tenant_name"])


# -------------------------------------------------
# STEP 1: SITE + LIST
# -------------------------------------------------
print("🔍 Finder site...")
site_id = client.get_site_id(SITE_NAME)

print("🔍 Finder liste...")
lists = client.get_lists(site_id)

target_list = next(l for l in lists if l["displayName"] == LIST_NAME)
list_id = target_list["id"]

print("✅ Liste fundet:", LIST_NAME)


# -------------------------------------------------
# STEP 2: MAPPING
# -------------------------------------------------
print("\n🔍 Genererer mapping...")

columns = client.get_columns(site_id, list_id)
mapping = generate_mapping(columns)

print("Mapping:")
print(json.dumps(mapping, indent=2))


# -------------------------------------------------
# STEP 3: TEST DATA (ALLE TYPER)
# -------------------------------------------------
print("\n🔍 Tester data...")

data = {
    # TEXT
    "title": "Test - Full mapping",

    # DATE (skal være ISO-format eller string)
    "startDato": "2026-06-01T10:00:00",

    # CHOICE (skal matche SharePoint valg)
    "status": "Aktiv",

    # LOOKUP (ID)
    "kunde": 1,

    # USER (kan være ID eller email afhængig opsætning)
    "ansvarlig": 23
}


print("Input data:")
print(json.dumps(data, indent=2))


# -------------------------------------------------
# STEP 4: MAP
# -------------------------------------------------
print("\n🔄 Mapper data...")

try:
    body = validate_and_map(data, mapping)

    print("✅ Mapping OK:")
    print(json.dumps(body, indent=2))

except Exception as e:
    print("❌ Mapping fejl:")
    print(e)
    exit()


# -------------------------------------------------
# STEP 5: CREATE ITEM
# -------------------------------------------------
print("\n📤 Opretter item...")

try:
    result = client.create_item(site_id, list_id, body)

    print("✅ Item oprettet!")
    print("ID:", result["id"])

except Exception as e:
    print("❌ Create fejlede:")
    print(e)


# -------------------------------------------------
# STEP 6: HENT OG VALIDÉR
# -------------------------------------------------
print("\n📥 Henter items...")

items = client.get_items(site_id, list_id)

print(f"✅ Antal items: {len(items)}")

latest = items[0]["fields"]

print("\n✅ Seneste item:")
print(json.dumps(latest, indent=2))