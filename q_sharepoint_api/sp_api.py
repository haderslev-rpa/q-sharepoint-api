from automation_server_client import AutomationServer, Credential

from q_sharepoint_api.sp_auth import SharePointAuth
from q_sharepoint_api.sp_client import SharePointClient


# -------------------------------------------------
# INIT (køres én gang)
# -------------------------------------------------
AutomationServer.from_environment()

_cred = Credential.get_credential("API_SHAREPOINT")
_cfg = _cred.data


# -------------------------------------------------
# SINGLETON (genbrug)
# -------------------------------------------------
_client = None


def get_client():
    global _client

    if _client:
        return _client

    auth = SharePointAuth(
        tenant_id=_cfg["tenant_id"],
        client_id=_cfg["client_id"],
        client_secret=_cred.password,
        scope=_cfg.get("scope", "https://graph.microsoft.com/.default")
    )

    _client = SharePointClient(auth, _cfg["tenant_name"])

    return _client