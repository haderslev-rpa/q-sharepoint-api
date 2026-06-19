# test_sp_drive.py

from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.sp_site_context import get_site


def run_test():
    print("🚀 Starter test...")

    client = get_client()  # objekt (konkret instans af klasse)

    # 🔧 KONFIGURATION
    site_name = "Automatisering"  # <-- ændr!
    base_path = "Test"

    # 1. Hent site
    site = get_site(site_name)
    site_id = site.site_id

    print("✅ Site ID:", site_id)

    # 2. Hent drive
    drive_id = client.get_drive_id(site_id)
    print("✅ Drive ID:", drive_id)

    # 3. Opret mappe
    folder = client.create_folder(
        drive_id,
        base_path,
        "MinTestMappe"
    )

    folder_id = folder["id"]
    print("✅ Mappe oprettet:", folder_id)

    # 4. Upload fil
    test_content = b"Hej fra test!"  # bytes (rå data)

    uploaded = client.upload_file(
        drive_id,
        f"{base_path}/MinTestMappe",
        "test.txt",
        test_content
    )

    print("✅ Fil uploadet:", uploaded["id"])

    # 5. Slet mappe (inkl. fil)
    client.delete_item(drive_id, folder_id)

    print("🗑️ Mappe slettet")

    print("✅ TEST FÆRDIG!")


if __name__ == "__main__":
    run_test()