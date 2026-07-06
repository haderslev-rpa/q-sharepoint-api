# q_sharepoint_api/sp_client.py

import requests
from urllib.parse import quote


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
    # GRAPH DRIVE (mapper + filer)
    # ---------------------------

    def get_drive_id(self, site_id):
        """
        Henter drive-id (dokumentbibliotek)
        """
        url = f"{self.base}/sites/{site_id}/drive"

        r = requests.get(
            url,
            headers=self.auth.graph_headers(),
            timeout=30
        )
        r.raise_for_status()

        return r.json()["id"]


    def create_folder(self, drive_id, parent_path, folder_name):
        """
        Opretter mappe i SharePoint

        parent_path: fx "Shared Documents/Test"
        folder_name: fx "MinMappe"
        """

        url = f"{self.base}/drives/{drive_id}/root:/{parent_path}:/children"

        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }

        r = requests.post(
            url,
            headers=self.auth.graph_headers(),
            json=data,
            timeout=30
        )
        r.raise_for_status()

        return r.json()  # objekt (konkret instans af mappe)


    def upload_file(self, drive_id, folder_path, file_name, content):
        """
        Uploader fil

        folder_path: fx "Shared Documents/Test/MinMappe"
        content: bytes (fil indhold)
        """

        url = (
            f"{self.base}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"
        )

        r = requests.put(
            url,
            headers=self.auth.graph_headers(),
            data=content,
            timeout=60
        )
        r.raise_for_status()

        return r.json()


    def delete_item(self, drive_id, item_id):
        """
        Sletter mappe eller fil via ID
        """

        url = f"{self.base}/drives/{drive_id}/items/{item_id}"

        r = requests.delete(
            url,
            headers=self.auth.graph_headers(),
            timeout=30
        )
        r.raise_for_status()


    def delete_file(self, drive_id, item_id):
        """
        Sletter fil (wrapper)
        """
        self.delete_item(drive_id, item_id)


    # ---------------------------
    # GRAPH EXCEL
    # ---------------------------

    
    def get_drive_item_id(self, site_id, file_path):
        """
        Finder item_id for en fil i SharePoint ud fra filsti.

        file_path:
            Sti inde i dokumentbiblioteket.
            Fx:
            "Test/min_fil.xlsx"
        """

        # -------------------------------------------------
        # Gør file_path ren
        # -------------------------------------------------

        # Fjern mellemrum før/efter
        file_path = file_path.strip()

        # Fjern start-slash hvis den findes
        file_path = file_path.lstrip("/")

        # Ret Windows backslash til normal slash
        file_path = file_path.replace("\\", "/")

        # URL-encode (gør sti sikker til Graph API)
        encoded_path = quote(file_path, safe="/")

        # -------------------------------------------------
        # Graph URL
        # -------------------------------------------------
        url = f"{self.base}/sites/{site_id}/drive/root:/{encoded_path}:"

        print("🔎 get_drive_item_id URL:")
        print(url)

        r = requests.get(
            url,
            headers=self.auth.graph_headers(),
            timeout=30
        )

        if r.status_code == 404:
            print("❌ Filen blev ikke fundet på denne sti:")
            print(file_path)
            print("Encoded path:")
            print(encoded_path)

        r.raise_for_status()

        return r.json()["id"]


    def create_excel_session(self, file_id):
        """
        Opretter session (hurtigere Excel operationer)
        """

        url = f"{self.base}/me/drive/items/{file_id}/workbook/createSession"

        data = {
            "persistChanges": True
        }

        r = requests.post(url, headers=self.auth.graph_headers(), json=data, timeout=30)
        r.raise_for_status()

        return r.json()["id"]


    def update_excel_range(self, file_id, sheet_name, cell_range, values, session_id=None):
        """
        Opdaterer Excel range (celler)

        values = [["A", "B", "C"]]  (liste (datastruktur))
        """

        url = (
            f"{self.base}/me/drive/items/{file_id}"
            f"/workbook/worksheets/{sheet_name}"
            f"/range(address='{cell_range}')"
        )

        headers = self.auth.graph_headers()

        # tilføj session hvis findes
        if session_id:
            headers["workbook-session-id"] = session_id

        data = {
            "values": values
        }

        r = requests.patch(url, headers=headers, json=data, timeout=30)
        r.raise_for_status()

        return r.json()

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
    

    # -------------------------------------------------
    # DOWNLOAD FILE TO MEMORY BY PATH
    # -------------------------------------------------
    def download_file_to_memory_by_path(
        self,
        site_id,
        file_path,
        save_dir=None
    ):
        """
        Henter en SharePoint-fil via fuld filsti og gemmer den i memory.

        site_id:
            SharePoint site id

        file_path:
            Sti inde i dokumentbiblioteket.
            Fx:
            "Distrikter/Vejliste over Haderslev kommune - Forebyggende Hje.xlsx"

        save_dir:
            Valgfri Linux-mappe.
            Hvis udfyldt, gemmes filen også på disk.
        """

        # -------------------------------------------------
        # Gør file_path API-sikker
        # -------------------------------------------------

        # Fjern evt. start-slash
        file_path = file_path.strip().lstrip("/")

        # Ret Windows backslash til normal slash
        file_path = file_path.replace("\\", "/")

        # URL-encode (gør sti sikker til API)
        encoded_file_path = quote(file_path, safe="/")

        # -------------------------------------------------
        # Find filen via path
        # -------------------------------------------------
        item_url = (
            f"{self.base}/sites/{site_id}"
            f"/drive/root:/{encoded_file_path}"
        )

        item_response = requests.get(
            item_url,
            headers=self.auth.graph_headers(),
            timeout=30
        )
        item_response.raise_for_status()

        item = item_response.json()

        item_id = item["id"]
        filename = item["name"]

        # -------------------------------------------------
        # Download fil-content
        # -------------------------------------------------
        content_url = (
            f"{self.base}/sites/{site_id}"
            f"/drive/items/{item_id}/content"
        )

        content_response = requests.get(
            content_url,
            headers=self.auth.graph_headers(),
            timeout=60
        )
        content_response.raise_for_status()

        file_bytes = content_response.content

        saved_path = None

        # -------------------------------------------------
        # Gem evt. på Linux
        # -------------------------------------------------
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)

            saved_path = os.path.join(save_dir, filename)

            with open(saved_path, "wb") as f:
                f.write(file_bytes)

        return {
            "filename": filename,
            "file_bytes": file_bytes,
            "saved_path": saved_path,
            "item": item
        }