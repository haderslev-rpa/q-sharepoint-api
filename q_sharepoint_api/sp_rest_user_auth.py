# sharepoint_rest_legacy/rest_user_auth.py

import time
import requests

from automation_server_client import AutomationServer, Credential


# -------------------------------------------------
# INIT (skal ske FØRST)
# -------------------------------------------------
AutomationServer.from_environment()  # miljø-init (krævet)


# -------------------------------------------------
# Credentials (hemmelige login-oplysninger)
# -------------------------------------------------

# Teknisk bruger (brugernavn + password)
_rest_user_cred = Credential.get_credential("API_COPILOT")

# SharePoint app-config (tenant, client-id)
_sp_cred = Credential.get_credential("API_SHAREPOINT")
_sp_cfg = _sp_cred.data


# -------------------------------------------------
# Token cache (gemmer token midlertidigt)
# -------------------------------------------------

_rest_token = None
_rest_token_expiry = 0
_TOKEN_BUFFER = 300  # sekunder (5 min)


# -------------------------------------------------
# PUBLIC FUNCTION
# -------------------------------------------------

def get_rest_user_token():
    """
    Henter USER-token til SharePoint REST
    Bruges KUN til liste-attachments
    """

    global _rest_token, _rest_token_expiry

    # Genbrug token hvis gyldigt
    if _rest_token and time.time() < (_rest_token_expiry - _TOKEN_BUFFER):
        return _rest_token

    token_url = (
        f"https://login.microsoftonline.com/"
        f"{_sp_cfg['tenant_id']}/oauth2/v2.0/token"
    )

    data = {
        "grant_type": "password",
        "client_id": _sp_cfg["client_id"],
        "client_secret": _sp_cred.password,
        "username": _rest_user_cred.username,
        "password": _rest_user_cred.password,
        "scope": f"https://{_sp_cfg['tenant_name']}.sharepoint.com/.default"
    }

    response = requests.post(token_url, data=data, timeout=30)
    response.raise_for_status()

    token_data = response.json()

    _rest_token = token_data["access_token"]
    _rest_token_expiry = time.time() + int(token_data.get("expires_in", 3600))

    return _rest_token


def get_rest_headers():
    """
    Returnerer headers til SharePoint REST
    """
    return {
        "Authorization": f"Bearer {get_rest_user_token()}",
        "Accept": "application/json;odata=verbose"
    }