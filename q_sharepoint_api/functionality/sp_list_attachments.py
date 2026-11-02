"""
Funktioner til vedhæftninger på SharePoint-liste-items.

Funktionerne bruger SharePoint REST API, fordi listeattachments
håndteres gennem listens UI-navn.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import requests

from q_sharepoint_api.sp_rest_user_auth import (
    get_rest_headers,
)
from q_sharepoint_api.sp_site_context import (
    get_site,
)


# ------------------------------------------------------------
# VALIDERING
# ------------------------------------------------------------

def _require_text(
    value: object,
    field_name: str,
) -> str:
    """
    Validerer en obligatorisk tekstværdi.

    Returnerer:
        Værdien som trimmet tekst.
    """
    text = str(value or "").strip()

    if not text:
        raise ValueError(
            f"{field_name} må ikke være tom."
        )

    return text


def _validate_file_name(
    file_name: str,
) -> str:
    """
    Validerer filnavnet til en SharePoint-vedhæftning.

    Filnavnet må kun være et filnavn og må ikke indeholde
    en lokal eller relativ mappesti.

    Returnerer:
        Det validerede filnavn.
    """
    validated_file_name = _require_text(
        file_name,
        "file_name",
    )

    invalid_characters = {
        "/",
        "\\",
        ":",
        "*",
        "?",
        '"',
        "<",
        ">",
        "|",
    }

    found_invalid_characters = sorted(
        character
        for character in invalid_characters
        if character in validated_file_name
    )

    if found_invalid_characters:
        raise ValueError(
            "file_name indeholder ugyldige tegn: "
            f"{found_invalid_characters}"
        )

    return validated_file_name


def _validate_file_bytes(
    file_bytes: bytes,
) -> bytes:
    """
    Validerer filindhold fra memory.

    Returnerer:
        De validerede bytes.
    """
    if not isinstance(file_bytes, bytes):
        raise TypeError(
            "file_bytes skal være bytes."
        )

    if not file_bytes:
        raise ValueError(
            "file_bytes må ikke være tom."
        )

    return file_bytes


def _prepare_attachment(
    *,
    file_path: str | os.PathLike[str] | None,
    file_name: str | None,
    file_bytes: bytes | None,
) -> tuple[str, bytes]:
    """
    Klargør vedhæftningen fra enten disk eller memory.

    Der skal vælges præcis én af disse inputmetoder:

    1. file_path
    2. file_bytes sammen med file_name

    Returnerer:
        Tuple med:
            - filnavn
            - filindhold som bytes
    """
    has_file_path = file_path is not None
    has_file_bytes = file_bytes is not None
    has_file_name = file_name is not None

    if has_file_path and has_file_bytes:
        raise ValueError(
            "Angiv enten file_path eller file_bytes, "
            "ikke begge dele."
        )

    if not has_file_path and not has_file_bytes:
        raise ValueError(
            "Angiv enten file_path eller file_bytes."
        )

    if has_file_path:
        if has_file_name:
            raise ValueError(
                "file_name må ikke angives sammen med file_path. "
                "Filnavnet udledes automatisk fra filstien."
            )

        path = Path(file_path)

        if not path.is_file():
            raise FileNotFoundError(
                f"Filen findes ikke: {path}"
            )

        validated_file_name = _validate_file_name(
            path.name
        )

        content = path.read_bytes()

        if not content:
            raise ValueError(
                f"Filen er tom: {path}"
            )

        return (
            validated_file_name,
            content,
        )

    if not has_file_name:
        raise ValueError(
            "file_name skal angives, når file_bytes bruges."
        )

    validated_file_name = _validate_file_name(
        file_name
    )

    validated_file_bytes = _validate_file_bytes(
        file_bytes
    )

    return (
        validated_file_name,
        validated_file_bytes,
    )


# ------------------------------------------------------------
# HENT ATTACHMENTS
# ------------------------------------------------------------

def list_attachments_for_item(
    site_name: str,
    list_name: str,
    item_id: int | str,
) -> list[dict[str, Any]]:
    """
    Henter alle vedhæftninger for et SharePoint-liste-item.

    SharePoint REST kræver listens UI-navn.

    Returnerer:
        En liste med SharePoints attachment-oplysninger.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )

    validated_list_name = _require_text(
        list_name,
        "list_name",
    )

    validated_item_id = _require_text(
        item_id,
        "item_id",
    )

    site = get_site(
        validated_site_name
    )

    escaped_list_name = (
        validated_list_name.replace(
            "'",
            "''",
        )
    )

    url = (
        f"{site.site_url}"
        f"/_api/web/lists/getbytitle('{escaped_list_name}')"
        f"/items({validated_item_id})/AttachmentFiles"
    )

    response = requests.get(
        url,
        headers=get_rest_headers(),
        timeout=30,
    )

    response.raise_for_status()

    return response.json()["d"]["results"]

def download_attachment_to_memory(
    site_name: str,
    attachment: dict[str, Any],
) -> dict[str, Any]:
    """
    Downloader én SharePoint-listeattachment direkte til memory.

    Input:
        site_name:
            SharePoint-sitets navn.

        attachment:
            Én attachment-dictionary fra
            list_attachments_for_item().

    Returnerer:
        {
            "file_name": "samtykke.pdf",
            "file_bytes": b"...",
            "content_type": "application/pdf",
            "size": 12345,
            "server_relative_url": "...",
            "attachment": {...},
        }

    Funktionen gemmer ikke filen på disk.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )

    if not isinstance(attachment, dict):
        raise TypeError(
            "attachment skal være en dictionary."
        )

    file_name = _validate_file_name(
        attachment.get("FileName")
    )

    server_relative_url = _require_text(
        attachment.get("ServerRelativeUrl"),
        "attachment['ServerRelativeUrl']",
    )

    site = get_site(
        validated_site_name
    )

    escaped_server_relative_url = (
        server_relative_url.replace(
            "'",
            "''",
        )
    )

    download_url = (
        f"{site.site_url}"
        f"/_api/web"
        f"/GetFileByServerRelativeUrl("
        f"'{escaped_server_relative_url}'"
        f")/$value"
    )

    response = requests.get(
        download_url,
        headers=get_rest_headers(),
        timeout=60,
    )

    response.raise_for_status()

    file_bytes = response.content

    if not isinstance(file_bytes, bytes):
        raise TypeError(
            "SharePoint returnerede ikke filindholdet som bytes."
        )

    if not file_bytes:
        raise ValueError(
            f"SharePoint-vedhæftningen {file_name!r} var tom."
        )

    return {
        "file_name": file_name,
        "file_bytes": file_bytes,
        "content_type": str(
            response.headers.get(
                "Content-Type",
                "",
            )
        ).strip(),
        "size": len(file_bytes),
        "server_relative_url": server_relative_url,
        "attachment": attachment,
    }


def get_attachments_for_item(
    site_name: str,
    list_name: str,
    item_id: int | str,
    *,
    include_file_bytes: bool = False,
) -> list[dict[str, Any]]:
    """
    Henter SharePoint-listeattachments for ét item.

    Med include_file_bytes=False returneres metadata uden
    filindhold:

        [
            {
                "file_name": "samtykke.pdf",
                "server_relative_url": "...",
                "attachment": {...},
            }
        ]

    Med include_file_bytes=True downloades alle attachments
    direkte til memory:

        [
            {
                "file_name": "samtykke.pdf",
                "file_bytes": b"...",
                "content_type": "application/pdf",
                "size": 12345,
                "server_relative_url": "...",
                "attachment": {...},
            }
        ]

    Hvis itemet ikke har attachments, returneres en tom liste.

    Funktionen gemmer ikke filer på disk.
    """
    if not isinstance(
        include_file_bytes,
        bool,
    ):
        raise TypeError(
            "include_file_bytes skal være bool."
        )

    attachments = list_attachments_for_item(
        site_name=site_name,
        list_name=list_name,
        item_id=item_id,
    )

    if not isinstance(attachments, list):
        raise TypeError(
            "SharePoint returnerede ikke attachments som en liste."
        )

    results: list[dict[str, Any]] = []

    for attachment in attachments:
        if not isinstance(attachment, dict):
            raise TypeError(
                "SharePoint returnerede en attachment, "
                "som ikke er en dictionary."
            )

        if include_file_bytes:
            result = download_attachment_to_memory(
                site_name=site_name,
                attachment=attachment,
            )
        else:
            file_name = _validate_file_name(
                attachment.get("FileName")
            )

            server_relative_url = _require_text(
                attachment.get(
                    "ServerRelativeUrl"
                ),
                (
                    "attachment"
                    "['ServerRelativeUrl']"
                ),
            )

            result = {
                "file_name": file_name,
                "server_relative_url":
                    server_relative_url,
                "attachment": attachment,
            }

        results.append(result)

    return results


# ------------------------------------------------------------
# DOWNLOAD ATTACHMENT
# ------------------------------------------------------------

def download_attachment(
    site_name: str,
    attachment: dict[str, Any],
    target_folder: str | os.PathLike[str],
) -> str:
    """
    Downloader én SharePoint-vedhæftning og gemmer den lokalt.

    Returnerer:
        Den lokale filsti som tekst.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )

    if not isinstance(attachment, dict):
        raise TypeError(
            "attachment skal være en dictionary."
        )

    file_name = _validate_file_name(
        attachment.get("FileName")
    )

    server_relative_url = _require_text(
        attachment.get("ServerRelativeUrl"),
        "attachment['ServerRelativeUrl']",
    )

    site = get_site(
        validated_site_name
    )

    escaped_server_relative_url = (
        server_relative_url.replace(
            "'",
            "''",
        )
    )

    download_url = (
        f"{site.site_url}"
        f"/_api/web"
        f"/GetFileByServerRelativeUrl("
        f"'{escaped_server_relative_url}'"
        f")/$value"
    )

    target_path = Path(
        target_folder
    )

    target_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    local_path = (
        target_path
        / file_name
    )

    response = requests.get(
        download_url,
        headers=get_rest_headers(),
        timeout=60,
    )

    response.raise_for_status()

    local_path.write_bytes(
        response.content
    )

    return str(local_path)


# ------------------------------------------------------------
# UPLOAD ATTACHMENT
# ------------------------------------------------------------

def add_attachment_to_item(
    site_name: str,
    list_name: str,
    item_id: int | str,
    file_path: str | os.PathLike[str] | None = None,
    *,
    file_name: str | None = None,
    file_bytes: bytes | None = None,
) -> dict[str, Any]:
    """
    Uploader en vedhæftning til et SharePoint-liste-item.

    Vælg præcis én inputmetode.

    Upload fra en lokal fil:

        add_attachment_to_item(
            site_name="Automatisering",
            list_name="Min liste",
            item_id=1,
            file_path="/tmp/dokument.pdf",
        )

    Upload direkte fra memory:

        add_attachment_to_item(
            site_name="Automatisering",
            list_name="Min liste",
            item_id=1,
            file_name="dokument.pdf",
            file_bytes=pdf_bytes,
        )

    Ved file_path udledes filnavnet automatisk fra stien.

    Ved file_bytes skal file_name angives.

    Returnerer:
        SharePoints JSON-svar som dictionary.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )

    validated_list_name = _require_text(
        list_name,
        "list_name",
    )

    validated_item_id = _require_text(
        item_id,
        "item_id",
    )

    (
        validated_file_name,
        validated_file_bytes,
    ) = _prepare_attachment(
        file_path=file_path,
        file_name=file_name,
        file_bytes=file_bytes,
    )

    site = get_site(
        validated_site_name
    )

    escaped_list_name = (
        validated_list_name.replace(
            "'",
            "''",
        )
    )

    escaped_file_name = (
        validated_file_name.replace(
            "'",
            "''",
        )
    )

    upload_url = (
        f"{site.site_url}"
        f"/_api/web/lists/getbytitle('{escaped_list_name}')"
        f"/items({validated_item_id})/AttachmentFiles/add"
        f"(FileName='{escaped_file_name}')"
    )

    headers = get_rest_headers()
    headers["Content-Type"] = (
        "application/octet-stream"
    )

    response = requests.post(
        upload_url,
        headers=headers,
        data=validated_file_bytes,
        timeout=60,
    )

    response.raise_for_status()

    return response.json()