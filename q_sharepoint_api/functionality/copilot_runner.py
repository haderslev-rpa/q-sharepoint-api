from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.copilot_auth import get_user_token_for_copilot

from q_sharepoint_api.sp_temp_files import (
    upload_temp_files,
    delete_temp_folder
)

from q_sharepoint_api.copilot_builder import build_body
from q_sharepoint_api.copilot_api import CopilotAPI


# -------------------------------------------------
# KONFIGURATION (fast)
# -------------------------------------------------
BASE_PATH = "/Robot - Copilot/Temp"
DEFAULT_SITE = "Automatisering"


# -------------------------------------------------
# INTERN: extract tekst
# -------------------------------------------------
def _extract_simple_output(messages):
    texts = []

    for m in messages:
        content = m.get("content") or m.get("text")

        if isinstance(content, list):
            for c in content:
                texts.append(c.get("text", ""))

        elif isinstance(content, str):
            texts.append(content)

    return "\n".join([t for t in texts if t])


# -------------------------------------------------
# MAIN FUNCTION
# -------------------------------------------------
def run_copilot(
    prompt,
    files=None,
    context=None,
    include_citations=True,
    simple_output=True,
    site_name=DEFAULT_SITE   # ✅ DEFAULT HER
):
    """
    🎯 Simpel Copilot wrapper (CURA-style)

    --------------------------------
    ✅ SIMPEL BRUG (99% cases)
    --------------------------------
    result = run_copilot(
        prompt="Hvad er 2+2?"
    )

    print(result)


    --------------------------------
    ✅ MED FILER
    --------------------------------
    result = run_copilot(
        prompt="Hvad står der i dokumentet?",
        files=[
            {"path": r"C:\\Temp\\test.pdf", "name": "test.pdf"}
        ]
    )

    print(result)


    --------------------------------
    ✅ ANDRE SITE (override)
    --------------------------------
    result = run_copilot(
        prompt="Test",
        site_name="FS_Oekonomiafdelingen"
    )


    --------------------------------
    PARAMETRE
    --------------------------------
    prompt:
        Spørgsmål til Copilot

    files:
        filer der uploades midlertidigt

    context:
        ekstra tekst (additionalContext)

    include_citations:
        om Copilot skal give kilder

    simple_output:
        True → kun tekst
        False → raw response

    site_name:
        Default = "Automatisering"
    """

    sp_client = get_client()
    token = get_user_token_for_copilot()

    drive_id = None
    folder_id = None

    try:
        file_urls = []

        # -------------------------------------------------
        # STEP 1: Upload filer (hvis der er nogen)
        # -------------------------------------------------
        if files:
            drive_id, folder_id, file_urls = upload_temp_files(
                sp_client,
                site_name,
                BASE_PATH,
                files
            )

        print("📎 File URLs:", file_urls)

        # -------------------------------------------------
        # STEP 2: Build body
        # -------------------------------------------------
        body = build_body(
            prompt=prompt,
            file_urls=file_urls,
            context=context,
            include_citations=include_citations
        )

        print("📦 Body klar")

        # -------------------------------------------------
        # STEP 3: Start Copilot
        # -------------------------------------------------
        copilot = CopilotAPI(token)

        convo = copilot.start_conversation()
        conversation_id = convo["id"]

        print("💬 Conversation ID:", conversation_id)

        # -------------------------------------------------
        # STEP 4: Send prompt
        # -------------------------------------------------
        response = copilot.send_message(conversation_id, body)

        messages = response.get("messages", [])

        # -------------------------------------------------
        # STEP 5: Output
        # -------------------------------------------------
        if simple_output:
            return _extract_simple_output(messages)

        return messages

    except Exception as e:
        print("❌ Fejl:", str(e))
        raise

    finally:
        # -------------------------------------------------
        # STEP 6: Cleanup (ALTID)
        # -------------------------------------------------
        if drive_id and folder_id:
            print("🧹 Cleanup starter...")
            try:
                delete_temp_folder(sp_client, drive_id, folder_id)
                print("✅ Temp mappe slettet")
            except Exception as cleanup_error:
                print("⚠️ Cleanup fejlede:", cleanup_error)