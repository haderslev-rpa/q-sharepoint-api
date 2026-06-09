# copilot_auth.py
import requests
from automation_server_client import Credential


_copilot_cred = Credential.get_credential("API_COPILOT")
_sp_cred = Credential.get_credential("API_SHAREPOINT")
_cfg = _sp_cred.data


def get_user_token_for_copilot():
    """
    Henter USER-token (delegated login)
    Bruges KUN til Copilot Graph
    """

    url = f"https://login.microsoftonline.com/{_cfg['tenant_id']}/oauth2/v2.0/token"

    data = {
        "grant_type": "password",
        "client_id": _cfg["client_id"],
        "client_secret": _sp_cred.password,
        "username": _copilot_cred.username,
        "password": _copilot_cred.password,
        "scope": "openid offline_access https://graph.microsoft.com/.default"
    }

    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()

    return r.json()["access_token"]