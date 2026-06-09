# q_sharepoint_api/site_context.py

from automation_server_client import AutomationServer, Credential
from q_sharepoint_api.sp_api import get_client


# INIT (miljø-opsætning)
AutomationServer.from_environment()


_sp_cred = Credential.get_credential("API_SHAREPOINT")
_sp_cfg = _sp_cred.data


class SiteContext:
    """
    SharePoint site-context (samlet site-info)
    """

    def __init__(self, site_name):
        self.site_name = site_name
        self.tenant_name = _sp_cfg["tenant_name"]

        self.site_url = (
            f"https://{self.tenant_name}.sharepoint.com/sites/{site_name}"
        )

        self.client = get_client()  # Graph-klient (singleton)
        self._site_id = None

    @property
    def site_id(self):
        """Graph site-id (cachet)"""
        if not self._site_id:
            self._site_id = self.client.get_site_id(self.site_name)
        return self._site_id


# ✅ ÉN-LINJES HELPER
def get_site(site_name):
    """
    Returnerer SiteContext (én linje)
    """
    return SiteContext(site_name)