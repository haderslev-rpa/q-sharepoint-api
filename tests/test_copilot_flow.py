from q_sharepoint_api.functionality.copilot_runner import run_copilot


files = [
    {
        "path": r"C:\Temp\test.pdf",
        "name": "test.pdf"
    }
]


result = run_copilot(
    site_name="Automatisering",
    base_path="/Robot - Copilot/Temp",
    prompt="Hvad står der i dokumentet?",
    files=files,
    context="Dette er en test"
)


print("\n✅ RESULTAT:")
for r in result:
    print(r)
