# tests/test_auth_smoke.py

"""
SMOKE TEST – AUTH & RETTIGHEDER

Tester:
1) Copilot (USER)
2) Microsoft Graph (APP)
3) SharePoint REST (USER)
"""

import requests
from automation_server_client import AutomationServer

from q_sharepoint_api.copilot_runner import run_copilot
from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.sp_rest_user_auth import get_rest_headers
from q_sharepoint_api.sp_site_context import get_site


# -------------------------------------------------
# INIT (miljø)
# -------------------------------------------------
AutomationServer.from_environment()

print("\n🚀 STARTER AUTH SMOKE TEST\n")


# -------------------------------------------------
# KONFIG
# -------------------------------------------------
SITE_NAME = "Automatisering"


# -------------------------------------------------
# TEST 1: COPILOT (USER)
# -------------------------------------------------
def test_copilot_user():
    print("🤖 Tester Copilot (USER)...")

    result = run_copilot(
        prompt="Svar kun med: Copilot auth OK ✅"
    )

    assert "Copilot" in result
    print("✅ Copilot USER auth virker")


# -------------------------------------------------
# TEST 2: GRAPH (APP)
# -------------------------------------------------
def test_graph_app():
    print("\n🔍 Tester Microsoft Graph (APP)...")

    client = get_client()  # Graph client (app)
    site_id = client.get_site_id(SITE_NAME)

    print("✅ Graph APP auth virker – Site ID:", site_id)


# -------------------------------------------------
# TEST 3: SHAREPOINT REST (USER)
# -------------------------------------------------
def test_rest_user():
    print("\n📎 Tester SharePoint REST (USER)...")

    site = get_site(SITE_NAME)

    url = f"{site.site_url}/_api/web"

    response = requests.get(
        url,
        headers=get_rest_headers(),
        timeout=30
    )
    response.raise_for_status()

    print("✅ SharePoint REST USER auth virker")


# -------------------------------------------------
# MAIN
# -------------------------------------------------
if __name__ == "__main__":
    try:
        test_copilot_user()
        test_graph_app()
        test_rest_user()

        print("\n🎉 AUTH SMOKE TEST BESTÅET")

    except Exception as e:
        print("\n❌ AUTH SMOKE TEST FEJLEDE")
        print(str(e))
        raise