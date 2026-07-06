from q_sharepoint_api.sp_api import get_client


def test_download_memory():
    """
    Tester download af SharePoint-fil til memory (RAM).
    Printer hvert trin, så vi kan se hvor det evt. fejler.
    """

    # -------------------------------------------------
    # INPUT
    # -------------------------------------------------
    site_name = "Automatisering"
    file_path = "Test/Vejliste over Haderslev kommune - Forebyggende Hjemmebesøg (1).xlsx"

    print("\n==============================")
    print("🔎 TEST: SharePoint fil til memory")
    print("==============================")

    print("\n📌 Input:")
    print("Site name:", site_name)
    print("File path:", file_path)

    # -------------------------------------------------
    # CLIENT
    # -------------------------------------------------
    print("\n1️⃣ Opretter SharePoint client...")

    try:
        client = get_client()
        print("✅ Client oprettet:", type(client))
    except Exception as e:
        print("❌ Fejl ved get_client()")
        print("Fejl:", repr(e))
        raise

    # -------------------------------------------------
    # SITE ID
    # -------------------------------------------------
    print("\n2️⃣ Finder site_id...")

    try:
        site_id = client.get_site_id(site_name)
        print("✅ site_id fundet:")
        print(site_id)
    except Exception as e:
        print("❌ Fejl ved get_site_id()")
        print("Site name brugt:", site_name)
        print("Fejl:", repr(e))
        print("\n💡 Hvis denne fejler, er site_name sandsynligvis forkert.")
        raise

    # -------------------------------------------------
    # DRIVE ID
    # -------------------------------------------------
    print("\n3️⃣ Finder drive_id...")

    try:
        drive_id = client.get_drive_id(site_id)
        print("✅ drive_id fundet:")
        print(drive_id)
    except Exception as e:
        print("❌ Fejl ved get_drive_id()")
        print("site_id brugt:", site_id)
        print("Fejl:", repr(e))
        raise

    # -------------------------------------------------
    # ITEM ID
    # -------------------------------------------------
    print("\n4️⃣ Finder filens item_id...")

    try:
        item_id = client.get_drive_item_id(site_id, file_path)
        print("✅ item_id fundet:")
        print(item_id)
    except Exception as e:
        print("❌ Fejl ved get_drive_item_id()")
        print("file_path brugt:", file_path)
        print("Fejl:", repr(e))
        print("\n💡 Hvis denne fejler, er file_path sandsynligvis forkert.")
        print("💡 Husk: Når du bruger /drive/root:/..., skal stien normalt være inde i dokumentbiblioteket.")
        print("💡 Eksempel: 'Test/fil.xlsx' og ikke nødvendigvis 'Dokumenter/Test/fil.xlsx'.")
        raise

    # -------------------------------------------------
    # DOWNLOAD FILE
    # -------------------------------------------------
    print("\n5️⃣ Downloader fil til memory...")

    try:
        result = client.download_file_to_memory_by_path(
            site_id=site_id,
            file_path=file_path,
            save_dir=None
        )

        filename = result["filename"]
        file_bytes = result["file_bytes"]

        print("✅ Fil hentet i memory")
        print("Filnavn:", filename)
        print("Bytes størrelse:", len(file_bytes))

    except Exception as e:
        print("❌ Fejl ved download_file()")
        print("drive_id brugt:", drive_id)
        print("item_id brugt:", item_id)
        print("Fejl:", repr(e))
        raise

    # -------------------------------------------------
    # MEMORY CHECK
    # -------------------------------------------------
    print("\n6️⃣ Tjekker memory-resultat...")

    if not file_bytes:
        raise Exception("Filen blev hentet, men file_bytes er tom.")

    print("✅ file_bytes findes og er ikke tom")
    print("Type på file_bytes:", type(file_bytes))

    print("\n==============================")
    print("✅ TEST GENNEMFØRT")
    print("==============================")

    return filename, file_bytes


if __name__ == "__main__":
    test_download_memory()