# q_sharepoint_api/sp_listattachments.py

import os
import requests

from q_sharepoint_api.sp_rest_user_auth import get_rest_headers
from q_sharepoint_api.sp_site_context import get_site


def list_attachments_for_item(site_name, list_name, item_id):
    """
    Henter alle attachments for et SharePoint liste-item
    (REST kræver listens UI-navn)
    """

    site = get_site(site_name)  # SiteContext (samlet site-info)

    url = (
        f"{site.site_url}/_api/web/lists/getbytitle('{list_name}')"
        f"/items({item_id})/AttachmentFiles"
    )

    response = requests.get(
        url,
        headers=get_rest_headers(),
        timeout=30
    )
    response.raise_for_status()

    return response.json()["d"]["results"]


def download_attachment(site_name, attachment, target_folder):
    """
    Downloader ét attachment og gemmer lokalt på Linux
    """

    site = get_site(site_name)  # SiteContext (samlet site-info)

    file_name = attachment["FileName"]
    server_relative_url = attachment["ServerRelativeUrl"]

    # ✅ KORREKT REST download-endpoint
    download_url = (
        f"{site.site_url}/_api/web"
        f"/GetFileByServerRelativeUrl('{server_relative_url}')/$value"
    )

    os.makedirs(target_folder, exist_ok=True)
    local_path = os.path.join(target_folder, file_name)

    response = requests.get(
        download_url,
        headers=get_rest_headers(),
        timeout=60
    )
    response.raise_for_status()

    with open(local_path, "wb") as f:
        f.write(response.content)

    return local_path


def add_attachment_to_item(site_name, list_name, item_id, file_path):
    """
    Uploader en lokal fil som attachment på et SharePoint liste-item
    (REST kræver listens UI-navn)
    """

    site = get_site(site_name)  # SiteContext (samlet site-info)

    file_name = os.path.basename(file_path)

    # Læs filen som bytes (rå fil-data)
    with open(file_path, "rb") as f:
        file_bytes = f.read()

    upload_url = (
        f"{site.site_url}/_api/web/lists/getbytitle('{list_name}')"
        f"/items({item_id})/AttachmentFiles/add"
        f"(FileName='{file_name}')"
    )

    headers = get_rest_headers()
    headers["Content-Type"] = "application/octet-stream"

    response = requests.post(
        upload_url,
        headers=headers,
        data=file_bytes,
        timeout=60
    )
    response.raise_for_status()

    return response.json()
