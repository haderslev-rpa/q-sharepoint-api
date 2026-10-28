
import time

import requests
from automation_server_client import AutomationServer, Credential


# -------------------------------------------------
# KONFIGURATION
# -------------------------------------------------

REST_USER_CREDENTIAL_NAME = "DIRXHEL"
SHAREPOINT_CREDENTIAL_NAME = "API_SHAREPOINT"

USERNAME_DOMAIN = "@haderslev.dk"

TOKEN_TIMEOUT_SECONDS = 30
TOKEN_BUFFER_SECONDS = 300


# -------------------------------------------------
# INIT (skal ske FØRST)
# -------------------------------------------------

AutomationServer.from_environment()


# -------------------------------------------------
# Credentials (hemmelige login-oplysninger)
# -------------------------------------------------

# Teknisk bruger:
# - username: DIRXHEL
# - password: Brugerens Microsoft-adgangskode
_rest_user_cred = Credential.get_credential(
    REST_USER_CREDENTIAL_NAME
)

# SharePoint app-config:
# - data: tenant_id, tenant_name og client_id
# - password: app registrationens client secret
_sp_cred = Credential.get_credential(
    SHAREPOINT_CREDENTIAL_NAME
)

_sp_cfg = _sp_cred.data or {}


# -------------------------------------------------
# Token cache (gemmer token midlertidigt)
# -------------------------------------------------

_rest_token = None
_rest_token_expiry = 0


# -------------------------------------------------
# INTERNE HJÆLPEFUNKTIONER
# -------------------------------------------------

def _require_text(value, field_name):
    """
    Kontrollerer, at en almindelig tekstværdi findes.

    Returnerer:
        Tekstværdien uden mellemrum før og efter.
    """
    if value is None:
        raise ValueError(f"{field_name} mangler.")

    cleaned_value = str(value).strip()

    if not cleaned_value:
        raise ValueError(f"{field_name} må ikke være tom.")

    return cleaned_value


def _require_secret(value, field_name):
    """
    Kontrollerer, at et password eller en secret findes.

    Værdien behandles ikke med strip(), fordi den skal sendes
    præcis, som den er gemt i Automation Server.

    Returnerer:
        Den oprindelige hemmelige tekstværdi.
    """
    if not isinstance(value, str) or not value:
        raise ValueError(f"{field_name} mangler eller er tom.")

    return value


def _build_rest_username(username):
    """
    Bygger brugerens fulde Microsoft-login.

    Eksempel:
        DIRXHEL -> DIRXHEL@haderslev.dk

    Hvis brugernavnet allerede indeholder @, bliver domænet
    ikke tilføjet igen.

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


# -------------------------------------------------
# PUBLIC FUNCTIONS
# -------------------------------------------------

def get_rest_user_token():
    """
    Henter et USER-token til SharePoint REST.

    Funktionen bruges kun til SharePoint-listevedhæftninger.

    Returnerer:
        Microsoft access token som tekst.
    """
    global _rest_token, _rest_token_expiry

    # Genbrug token, hvis det stadig er gyldigt.
    if (
        _rest_token
        and time.time()
        < (_rest_token_expiry - TOKEN_BUFFER_SECONDS)
    ):
        return _rest_token

    tenant_id = _require_text(
        _sp_cfg.get("tenant_id"),
        "API_SHAREPOINT.data['tenant_id']",
    )

    tenant_name = _require_text(
        _sp_cfg.get("tenant_name"),
        "API_SHAREPOINT.data['tenant_name']",
    )

    client_id = _require_text(
        _sp_cfg.get("client_id"),
        "API_SHAREPOINT.data['client_id']",
    )

    client_secret = _require_secret(
        _sp_cred.password,
        "API_SHAREPOINT.password",
    )

    username = _build_rest_username(
        _rest_user_cred.username
    )

    password = _require_secret(
        _rest_user_cred.password,
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
        "scope": (
            f"https://{tenant_name}.sharepoint.com/.default"
        ),
    }

    try:
        response = requests.post(
            token_url,
            data=request_data,
            timeout=TOKEN_TIMEOUT_SECONDS,
        )
    except requests.RequestException as error:
        raise RuntimeError(
            "Kunne ikke kontakte Microsofts token-endpoint "
            "for SharePoint REST. "
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
            "Microsoft afviste SharePoint REST-login.\n"
            f"HTTP-status: {response.status_code}\n"
            f"Fejl: {microsoft_error}\n"
            f"Beskrivelse: {microsoft_description}\n\n"
            "Kontrollér især:\n"
            "1. API_SHAREPOINT.data['tenant_id'].\n"
            "2. API_SHAREPOINT.data['tenant_name'].\n"
            "3. API_SHAREPOINT.data['client_id'].\n"
            "4. API_SHAREPOINT.password som client secret.\n"
            "5. DIRXHEL.username og DIRXHEL.password.\n"
            f"Microsoft-brugernavnet sendes som: {username}"
        )

    try:
        token_data = response.json()
    except ValueError as error:
        raise RuntimeError(
            "Microsoft returnerede et svar, som ikke var gyldig JSON."
        ) from error

    access_token = token_data.get("access_token")

    if not access_token:
        raise RuntimeError(
            "Microsofts tokensvar indeholder ikke 'access_token'. "
            f"Returnerede felter: {sorted(token_data.keys())}"
        )

    _rest_token = access_token
    _rest_token_expiry = (
        time.time()
        + int(token_data.get("expires_in", 3600))
    )

    return _rest_token


def get_rest_headers():
    """
    Returnerer HTTP-headers til SharePoint REST.

    Returnerer:
        En dictionary med access token og ønsket JSON-format.
    """
    return {
        "Authorization": f"Bearer {get_rest_user_token()}",
        "Accept": "application/json;odata=verbose",
    }