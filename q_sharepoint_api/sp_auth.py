# sp_auth.py
import time
import requests

class SharePointAuth:
    def __init__(self, tenant_id, tenant_name, client_id, client_secret, scope):
        self.tenant_id = tenant_id
        self.tenant_name = tenant_name
        self.client_id = client_id
        self.client_secret = client_secret

        self.graph_scope = scope
        self.sp_scope = f"https://{tenant_name}.sharepoint.com/.default"

        self._graph_token = None
        self._graph_expiry = 0

        self._sp_token = None
        self._sp_expiry = 0

    # ---------------------------
    # GRAPH TOKEN (APP)
    # ---------------------------
    def _get_graph_token(self):
        if self._graph_token and time.time() < (self._graph_expiry - 300):
            return self._graph_token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": self.graph_scope
        }

        r = requests.post(url, data=data, timeout=30)
        r.raise_for_status()

        token = r.json()
        self._graph_token = token["access_token"]
        self._graph_expiry = time.time() + int(token.get("expires_in", 3600))

        return self._graph_token

    # ---------------------------
    # SHAREPOINT REST TOKEN (APP)
    # ---------------------------
    def _get_sp_token(self):
        if self._sp_token and time.time() < (self._sp_expiry - 300):
            return self._sp_token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": self.sp_scope
        }

        r = requests.post(url, data=data, timeout=30)
        r.raise_for_status()

        token = r.json()
        self._sp_token = token["access_token"]
        self._sp_expiry = time.time() + int(token.get("expires_in", 3600))

        return self._sp_token

    # ---------------------------
    # HEADERS
    # ---------------------------
    def graph_headers(self):
        return {
            "Authorization": f"Bearer {self._get_graph_token()}",
            "Content-Type": "application/json"
        }

    def rest_headers(self):
        return {
            "Authorization": f"Bearer {self._get_sp_token()}",
            "Accept": "application/json;odata=verbose"
        }