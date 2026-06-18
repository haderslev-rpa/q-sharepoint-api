from q_sharepoint_api.functionality.copilot_runner import run_copilot


# -------------------------------------------------
# TEST 1: HELT SIMPEL
# -------------------------------------------------
print("🚀 TEST 1: Simpel prompt\n")

result = run_copilot(
    prompt="Hvad er 2+2?"
)

print("✅ COPILOT SVAR:\n")
print(result)


# -------------------------------------------------
# TEST 2: Med context (optional)
# -------------------------------------------------
print("\n\n🚀 TEST 2: Med context\n")

result = run_copilot(
    prompt="Hvad er koden?",
    context="Koden er 12345"
)

print("✅ COPILOT SVAR:\n")
print(result)


# -------------------------------------------------
# TEST 3: Raw output (debug mode)
# -------------------------------------------------
print("\n\n🚀 TEST 3: Raw output\n")

result = run_copilot(
    prompt="Hvad er hovedstaden i Danmark?",
    simple_output=False
)

print("✅ RAW OUTPUT:\n")
print(result)


# -------------------------------------------------
# TEST 4: Override site (valgfri)
# -------------------------------------------------
print("\n\n🚀 TEST 4: Override site\n")

result = run_copilot(
    prompt="Sig hej",
    site_name="Automatisering"  # eller andet site
)

print("✅ COPILOT SVAR:\n")
print(result)