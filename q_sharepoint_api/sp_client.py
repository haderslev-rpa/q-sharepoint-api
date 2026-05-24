import requests

class SharePointClient:
    def __init__(self, auth, tenant_name):
        self.auth = auth
        self.tenant_name = tenant_name
        self.base = "https://graph.microsoft.com/v1.0"

    # ---------------------------
    # GRAPH
    # ---------------------------
    def get_site_id(self, site_name):
        url = f"{self.base}/sites/{self.tenant_name}.sharepoint.com:/sites/{site_name}"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["id"]

    def get_lists(self, site_id):
        url = f"{self.base}/sites/{site_id}/lists"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["value"]

    def get_items(self, site_id, list_id):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items?expand=fields"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["value"]

    def update_item(self, site_id, list_id, item_id, data):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        r = requests.patch(url, headers=self.auth.graph_headers(), json=data, timeout=30)
        r.raise_for_status()
        return r.json()

    def create_item(self, site_id, list_id, data):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items"
        payload = {"fields": data}
        r = requests.post(url, headers=self.auth.graph_headers(), json=payload, timeout=30)
        r.raise_for_status()
        return r.json()

    def get_columns(self, site_id, list_id):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/columns"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["value"]

    # ---------------------------
    # REST (attachments)
    # ---------------------------
    def get_attachments(self, site_name, list_title, item_id):
        site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{site_name}"

        url = f"{site_url}/_api/web/lists/getbytitle('{list_title}')/items({item_id})/AttachmentFiles"

        r = requests.get(url, headers=self.auth.rest_headers(), timeout=30)
        r.raise_for_status()

        return r.json()