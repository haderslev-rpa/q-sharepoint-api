def build_body(prompt, file_urls=None, context=None, include_citations=True):
    """
    Erstatter AL din JSON manipulation fra Blue Prism
    """

    return {
        "message": {
            "text": prompt
        },
        "locationHint": {
            "timeZone": "Europe/Copenhagen"
        },
        "contextualResources": {
            "files": [{"uri": u} for u in (file_urls or [])]
        },
        "additionalContext": (
            [{"text": context}] if context else []
        ),
        "modelRequest": {
            "includeCitations": include_citations
        }
    }