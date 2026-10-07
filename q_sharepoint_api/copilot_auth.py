import requests
from automation_server_client import Credential


TOKEN_TIMEOUT_SECONDS = 30

COPILOT_CREDENTIAL_NAME = "DIRXHEL"
SHAREPOINT_CREDENTIAL_NAME = "API_SHAREPOINT"

USERNAME_DOMAIN = "@haderslev.dk"

COPILOT_SCOPE = (
    "openid "
    "offline_access "
    "https://graph.microsoft.com/.default"
)


def _require_text(value, field_name):
    """
    Kontrollerer, at en tekstværdi findes.

    Returnerer:
        Teksten uden mellemrum før og efter.

    Rejser:
        ValueError, hvis værdien mangler eller er tom.
    """
    if value is None:
        raise ValueError(f"{field_name} mangler.")

    cleaned_value = str(value).strip()

    if not cleaned_value:
        raise ValueError(f"{field_name} må ikke være tom.")

    return cleaned_value


def _require_secret(value, field_name):
    """
    Kontrollerer, at en hemmelig værdi findes.

    Password og client secret bliver ikke behandlet med strip(),
    fordi værdien skal sendes præcis, som den er gemt.

    Returnerer:
        Den oprindelige hemmelige tekstværdi.

    Rejser:
        ValueError, hvis værdien mangler eller er tom.
    """
    if not isinstance(value, str) or not value:
        raise ValueError(f"{field_name} mangler eller er tom.")

    return value


def _build_copilot_username(username):
    """
    Bygger det fulde Microsoft-brugernavn.

    Eksempel:
        DIRXHEL bliver til DIRXHEL@haderslev.dk

    Hvis username allerede indeholder @, tilføjes domænet ikke igen.

    Returnerer:
        Det fulde Microsoft-brugernavn.
    """
    cleaned_username = _require_text(
        username,
        "DIRXHEL.username",
    )

    if "@" in cleaned_username:
        return cleaned_username

    return f"{cleaned_username}{USERNAME_DOMAIN}"


def get_user_token_for_copilot():
    """
    Henter et delegated user-token til Microsoft Copilot Graph.

    Credentials:

        API_SHAREPOINT:
            Data JSON:
                {
                    "tenant_id": "...",
                    "client_id": "..."
                }

            Password:
                Client secret til app registration.

        DIRXHEL:
            Username:
                Kort brugernavn, eksempelvis DIRXHEL.

            Password:
                Brugerens Microsoft-adgangskode.

    Brugernavnet DIRXHEL bliver automatisk lavet om til:
        DIRXHEL@haderslev.dk

    Returnerer:
        Microsoft access token som tekst.
    """
    copilot_credential = Credential.get_credential(
        COPILOT_CREDENTIAL_NAME
    )
    sharepoint_credential = Credential.get_credential(
        SHAREPOINT_CREDENTIAL_NAME
    )

    sharepoint_config = sharepoint_credential.data or {}

    tenant_id = _require_text(
        sharepoint_config.get("tenant_id"),
        "API_SHAREPOINT.data['tenant_id']",
    )

    client_id = _require_text(
        sharepoint_config.get("client_id"),
        "API_SHAREPOINT.data['client_id']",
    )

    client_secret = _require_secret(
        sharepoint_credential.password,
        "API_SHAREPOINT.password",
    )

    username = _build_copilot_username(
        copilot_credential.username
    )

    password = _require_secret(
        copilot_credential.password,
        "DIRXHEL.password",
    )

    token_url = (
        "https://login.microsoftonline.com/"
        f"{tenant_id}/oauth2/v2.0/token"
    )

    request_data = {
        "grant_type": "password",
        "client_id": client_id,
        "client_secret": client_secret,
        "username": username,
        "password": password,
        "scope": COPILOT_SCOPE,
    }

    try:
        response = requests.post(
            token_url,
            data=request_data,
            timeout=TOKEN_TIMEOUT_SECONDS,
        )
    except requests.RequestException as error:
        raise RuntimeError(
            "Kunne ikke kontakte Microsofts token-endpoint. "
            f"Teknisk fejl: {error}"
        ) from error

    if not response.ok:
        try:
            error_data = response.json()
        except ValueError:
            error_data = {}

        microsoft_error = error_data.get(
            "error",
            "ukendt Microsoft-fejl",
        )

        microsoft_description = error_data.get(
            "error_description",
            response.text
            or "Microsoft returnerede ingen fejlbeskrivelse.",
        )

        raise RuntimeError(
            "Microsoft afviste Copilot-login.\n"
            f"HTTP-status: {response.status_code}\n"
            f"Fejl: {microsoft_error}\n"
            f"Beskrivelse: {microsoft_description}\n\n"
            "Kontrollér især:\n"
            "1. API_SHAREPOINT.data['tenant_id'].\n"
            "2. API_SHAREPOINT.data['client_id'].\n"
            "3. API_SHAREPOINT.password, som skal være client secret.\n"
            "4. DIRXHEL.username, som eksempelvis skal være DIRXHEL.\n"
            "5. DIRXHEL.password, som skal være brugerens adgangskode.\n"
            f"Microsoft-brugernavnet sendes som: {username}"
        )

    try:
        token_data = response.json()
    except ValueError as error:
        raise RuntimeError(
            "Microsoft returnerede HTTP 200, "
            "men svaret var ikke gyldig JSON."
        ) from error

    access_token = token_data.get("access_token")

    if not access_token:
        raise RuntimeError(
            "Microsofts tokensvar indeholder ikke 'access_token'. "
            f"Returnerede felter: {sorted(token_data.keys())}"
        )

    return access_token