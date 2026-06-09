# q_sharepoint_api/sp_client.py

import requests


class SharePointClient:
    def __init__(self, auth, tenant_name):
        self.auth = auth                    # auth (token-hjælper)
        self.tenant_name = tenant_name      # tenant navn
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

    def get_list_id(self, site_id, list_name):
        for lst in self.get_lists(site_id):
            if lst["displayName"] == list_name:
                return lst["id"]
        raise Exception(f"Liste '{list_name}' findes ikke")

    def get_list_items_raw(self, site_id, list_id):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items?expand=fields"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["value"]

    def create_list_item(self, site_id, list_id, payload):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items"
        r = requests.post(url, headers=self.auth.graph_headers(), json=payload, timeout=30)
        r.raise_for_status()
        return r.json()

    def update_list_item(self, site_id, list_id, item_id, payload):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        r = requests.patch(url, headers=self.auth.graph_headers(), json=payload, timeout=30)
        r.raise_for_status()
        return r.json()

    def get_list_columns(self, site_id, list_id):
        url = f"{self.base}/sites/{site_id}/lists/{list_id}/columns"
        r = requests.get(url, headers=self.auth.graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()["value"]

    # ---------------------------
    # REST (attachments)
    # ---------------------------
    def get_attachments(self, site_name, list_title, item_id):
        """
        Henter attachments via SharePoint REST
        """

        site_url = f"https://{self.tenant_name}.sharepoint.com/sites/{site_name}"

        url = (
            f"{site_url}/_api/web/lists/getbytitle('{list_title}')"
            f"/items({item_id})/AttachmentFiles"
        )

        r = requests.get(url, headers=self.auth.rest_headers(), timeout=30)
        r.raise_for_status()

        return r.json()