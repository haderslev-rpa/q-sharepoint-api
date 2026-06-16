import json

from automation_server_client import AutomationServer, Credential
from src.q_sharepoint_api.sp_auth import SharePointAuth
from src.q_sharepoint_api.sp_client import SharePointClient
from src.q_sharepoint_api.sp_generator import generate_mapping
from src.q_sharepoint_api.sp_mapping import validate_and_map

# -------------------------------------------------
# INIT: Hent credentials fra Automation Server
# -------------------------------------------------
AutomationServer.from_environment()
cred = Credential.get_credential("API_SHAREPOINT")

cfg = cred.data

# -------------------------------------------------
# INPUT (kan senere komme fra RPA)
# -------------------------------------------------
SITE_NAME = "Automatisering"   # Navn på dit SharePoint site
LIST_NAME = "TestListe"        # Navn på din SharePoint liste
ITEM_ID = 1                    # ID på item du vil opdatere

# -------------------------------------------------
# AUTH + CLIENT (samme design som CURA)
# -------------------------------------------------
auth = SharePointAuth(
    tenant_id=cfg["tenant_id"],
    client_id=cfg["client_id"],
    client_secret=cred.password,
    scope=cfg.get("scope", "https://graph.microsoft.com/.default")
)

client = SharePointClient(auth, cfg["tenant_name"])


# -------------------------------------------------
# STEP 1: Find site_id (Graph kræver ID – ikke URL)
# -------------------------------------------------
print("🔍 Finder site...")
site_id = client.get_site_id(SITE_NAME)

# Example output:
# "haderslevdk.sharepoint.com,12345,67890"
print("SITE ID:", site_id)


# -------------------------------------------------
# STEP 2: Find liste (vi slår op ud fra navn)
# -------------------------------------------------
print("🔍 Finder lister...")
lists = client.get_lists(site_id)

# Vi finder den liste der matcher LIST_NAME
target_list = next(l for l in lists if l["displayName"] == LIST_NAME)
list_id = target_list["id"]

print("LIST ID:", list_id)


# -------------------------------------------------
# STEP 3: GENERATE MAPPING (💡 vigtig!)
# -------------------------------------------------
print("🔍 Genererer mapping...")

# Henter kolonner fra SharePoint
columns = client.get_columns(site_id, list_id)

# Genererer din JSON → SharePoint mapping
mapping = generate_mapping(columns)

print("Mapping preview:")
print(json.dumps(mapping, indent=2))


# -------------------------------------------------
# STEP 4: DATA (dette er din RPA JSON)
# -------------------------------------------------
print("🔍 Klargør data...")

data = {
    "title": "Test fra Python"
}

# 💡 dette er hvordan din robot skal levere data:
# {
#   "sharepoint": {
#       "title": ...
#   }
# }


# -------------------------------------------------
# STEP 5: MAP TIL SHAREPOINT FORMAT
# -------------------------------------------------
print("🔍 Mapper data...")

body = validate_and_map(data, mapping)

print("Body sendt til SharePoint:")
print(json.dumps(body, indent=2))


# -------------------------------------------------
# STEP 6: UPDATE ITEM
# -------------------------------------------------
print("🔍 Opdaterer item...")

try:
    result = client.update_item(site_id, list_id, ITEM_ID, body)
    print("✅ Update OK:", result)

except Exception as e:
    print("❌ Update fejlede (forventet hvis ingen rettigheder)")
    print(e)