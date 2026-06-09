import os
import datetime

from q_sharepoint_api.sp_site_context import get_site  # ✅ NY (lille)


def _generate_folder_name():
    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    machine = os.environ.get("COMPUTERNAME", "machine")
    return f"{now}({machine})"


def upload_temp_files(client, site_name, base_path, files):
    """
    Opret temp mappe → upload filer → returner URLs
    """

    # ✅ ENESTE ÆNDRING:
    site_id = get_site(site_name).site_id

    drive_id = client.get_drive_id(site_id)

    folder_name = _generate_folder_name()
    full_path = f"{base_path}/{folder_name}"

    print("📁 Opretter temp mappe:", full_path)

    folder = client.create_folder(drive_id, base_path, folder_name)
    folder_id = folder["id"]

    urls = []

    for f in files:
        print("⬆️ Upload:", f["name"])

        with open(f["path"], "rb") as file:
            uploaded = client.upload_file(
                drive_id,
                full_path,
                f["name"],
                file.read()
            )

        urls.append(uploaded["webUrl"])

    return drive_id, folder_id, urls


def delete_temp_folder(client, drive_id, folder_id):
    print("🗑️ Sletter temp mappe:", folder_id)
    client.delete_item(drive_id, folder_id)