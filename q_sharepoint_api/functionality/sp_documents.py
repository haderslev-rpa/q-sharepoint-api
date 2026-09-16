"""
Generelle funktioner til dokumentbiblioteker i SharePoint.

Filen kender ikke Automation Server-items eller Digital Post.
Funktionerne kan derfor genbruges af andre processer.

Funktioner:

- upload_document()
- download_document()
- get_documents()
- delete_document()
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any
from urllib.parse import quote

import requests

from q_sharepoint_api.sp_api import get_client
from q_sharepoint_api.sp_list_schema import get_list_schema


DEFAULT_PAGE_SIZE = 200
MAX_PAGE_SIZE = 999


def upload_document(
    *,
    site_name: str,
    library_name: str,
    file_name: str,
    content: bytes,
    fields: dict[str, Any] | None = None,
    conflict_behavior: str = "fail",
) -> dict[str, Any]:
    """
    Uploader ét dokument til et SharePoint-dokumentbibliotek.

    Dokumentet uploades direkte i roden af biblioteket.

    Parametre:
        site_name:
            SharePoint-sitets navn.
            Eksempel: "Automatisering".

        library_name:
            Dokumentbibliotekets viste navn.
            Eksempel: "Digitalpost".

        file_name:
            Filnavnet i SharePoint.
            Eksempel: "<dokument_id>.pdf".

        content:
            Filens binære indhold som bytes.

        fields:
            Valgfrie SharePoint-kolonner, som skal opdateres
            efter uploaden.

            Eksempel:

            {
                "dokument_id": "...",
                "forsendelses_id": "...",
            }

        conflict_behavior:
            Hvad Graph skal gøre, hvis filnavnet allerede findes.

            Tilladte værdier:
            - "fail"
            - "replace"
            - "rename"

            Standard er "fail".

    Output:
        En dictionary med blandt andet:

        {
            "drive_item_id": "...",
            "list_item_id": "...",
            "drive_id": "...",
            "file_name": "...",
            "web_url": "...",
            "size": 12345,
            "created": "...",
            "fields": {
                "dokument_id": "...",
                "forsendelses_id": "...",
            },
        }

    Fejl:
        Hvis upload eller opdatering af kolonner fejler,
        bliver der rejst en exception.

        Hvis filen er uploadet, men kolonnerne ikke kan opdateres,
        forsøger funktionen at slette filen igen.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )
    validated_library_name = _require_text(
        library_name,
        "library_name",
    )
    validated_file_name = _validate_file_name(file_name)
    validated_content = _validate_content(content)

    allowed_conflict_behaviors = {
        "fail",
        "replace",
        "rename",
    }

    if conflict_behavior not in allowed_conflict_behaviors:
        raise ValueError(
            "conflict_behavior skal være "
            "'fail', 'replace' eller 'rename'."
        )

    client = get_client()

    site_id = client.get_site_id(
        validated_site_name
    )

    library = _get_document_library(
        client=client,
        site_id=site_id,
        library_name=validated_library_name,
    )

    drive_id = library["drive_id"]
    list_id = library["list_id"]

    encoded_file_name = quote(
        validated_file_name,
        safe="",
    )

    upload_url = (
        f"{client.base}"
        f"/drives/{drive_id}"
        f"/root:/{encoded_file_name}:/content"
    )

    params = {
        "@microsoft.graph.conflictBehavior": (
            conflict_behavior
        )
    }

    headers = client.auth.graph_headers()
    headers["Content-Type"] = "application/octet-stream"

    response = requests.put(
        upload_url,
        headers=headers,
        params=params,
        data=validated_content,
        timeout=120,
    )
    response.raise_for_status()

    drive_item = response.json()
    drive_item_id = drive_item.get("id")

    if not drive_item_id:
        raise RuntimeError(
            "SharePoint-upload returnerede ikke et drive item-id."
        )

    try:
        list_item = _get_list_item_for_drive_item(
            client=client,
            drive_id=drive_id,
            drive_item_id=drive_item_id,
        )

        list_item_id = list_item.get("id")

        if not list_item_id:
            raise RuntimeError(
                "SharePoint returnerede ikke et list item-id."
            )

        if fields:
            api_fields = _map_ui_fields_to_api_fields(
                site_name=validated_site_name,
                library_name=validated_library_name,
                fields=fields,
            )

            _update_document_fields(
                client=client,
                site_id=site_id,
                list_id=list_id,
                list_item_id=list_item_id,
                fields=api_fields,
            )

        fresh_item = _get_document_by_drive_item_id(
            client=client,
            drive_id=drive_id,
            drive_item_id=drive_item_id,
        )

        return {
            "drive_item_id": drive_item_id,
            "list_item_id": str(list_item_id),
            "drive_id": drive_id,
            "list_id": list_id,
            "file_name": fresh_item.get(
                "name",
                validated_file_name,
            ),
            "web_url": fresh_item.get("webUrl"),
            "size": fresh_item.get("size"),
            "created": fresh_item.get(
                "createdDateTime"
            ),
            "e_tag": fresh_item.get("eTag"),
            "fields": dict(fields or {}),
        }

    except Exception:
        # Uploaden er gennemført, men resten af handlingen fejlede.
        # Forsøg at fjerne den delvist oprettede fil.
        try:
            delete_document(
                site_name=validated_site_name,
                library_name=validated_library_name,
                drive_item_id=drive_item_id,
            )
        except Exception:
            # Den oprindelige exception skal bevares.
            # En eventuel efterladt fil findes senere af cleanup.
            pass

        raise


def download_document(
    *,
    site_name: str,
    library_name: str,
    drive_item_id: str | None = None,
    dokument_id: str | None = None,
) -> bytes:
    """
    Henter et SharePoint-dokument direkte til memory.

    Der skal angives præcis én af følgende:

    - drive_item_id
    - dokument_id

    drive_item_id er hurtigst, fordi SharePoint-filen kan
    hentes direkte uden et søgekald.

    dokument_id foretager først et opslag i dokumentbiblioteket.

    Output:
        Filens binære indhold som bytes.

    Eksempel:

        pdf_content = download_document(
            site_name="Automatisering",
            library_name="Digitalpost",
            drive_item_id="01ABC...",
        )
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )
    validated_library_name = _require_text(
        library_name,
        "library_name",
    )

    supplied_identifiers = sum(
        value is not None
        for value in (
            drive_item_id,
            dokument_id,
        )
    )

    if supplied_identifiers != 1:
        raise ValueError(
            "Angiv præcis én af drive_item_id eller dokument_id."
        )

    client = get_client()

    site_id = client.get_site_id(
        validated_site_name
    )

    library = _get_document_library(
        client=client,
        site_id=site_id,
        library_name=validated_library_name,
    )

    drive_id = library["drive_id"]

    if dokument_id is not None:
        result = get_documents(
            site_name=validated_site_name,
            library_name=validated_library_name,
            filters={
                "dokument_id": dokument_id,
            },
            page_size=2,
            max_items=2,
        )

        if result["count"] == 0:
            raise FileNotFoundError(
                f"Dokument med dokument_id "
                f"{dokument_id!r} blev ikke fundet."
            )

        if result["count"] > 1:
            raise RuntimeError(
                f"Der blev fundet flere dokumenter med "
                f"dokument_id {dokument_id!r}."
            )

        drive_item_id = result["documents"][0][
            "drive_item_id"
        ]

    validated_drive_item_id = _require_text(
        drive_item_id,
        "drive_item_id",
    )

    download_url = (
        f"{client.base}"
        f"/drives/{drive_id}"
        f"/items/{validated_drive_item_id}"
        f"/content"
    )

    response = requests.get(
        download_url,
        headers=client.auth.graph_headers(),
        timeout=120,
        allow_redirects=True,
    )
    response.raise_for_status()

    return response.content


def get_documents(
    *,
    site_name: str,
    library_name: str,
    filters: dict[str, Any] | None = None,
    created_before: datetime | str | None = None,
    created_after: datetime | str | None = None,
    page_size: int = DEFAULT_PAGE_SIZE,
    max_items: int | None = None,
) -> dict[str, Any]:
    """
    Henter dokumenter fra et SharePoint-dokumentbibliotek.

    Funktionen er fleksibel og kan bruges til:

    - opslag på dokument_id
    - opslag på forsendelses_id
    - oprydning efter oprettelsesdato
    - fremtidige SharePoint-kolonner

    Parametre:
        filters:
            Dictionary med præcise kolonnefiltre.

            Eksempel:

            {
                "forsendelses_id": "...",
            }

        created_before:
            Hent kun dokumenter oprettet før tidspunktet.

        created_after:
            Hent kun dokumenter oprettet efter tidspunktet.

        page_size:
            Antal elementer pr. Graph-side.
            Standard er 200.

        max_items:
            Valgfrit maksimum for det samlede resultat.

            None betyder, at alle sider hentes.

    Output:
        En dictionary:

        {
            "exists": True,
            "count": 2,
            "documents": [
                {
                    "drive_item_id": "...",
                    "list_item_id": "...",
                    "file_name": "...",
                    "created": "...",
                    "modified": "...",
                    "size": 12345,
                    "dokument_id": "...",
                    "forsendelses_id": "...",
                }
            ],
        }

    Vigtigt:
        Funktionen anvender pagination og henter derfor ikke
        nødvendigvis kun den første Graph-side.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )
    validated_library_name = _require_text(
        library_name,
        "library_name",
    )
    validated_page_size = _validate_page_size(page_size)
    validated_max_items = _validate_max_items(max_items)

    client = get_client()

    site_id = client.get_site_id(
        validated_site_name
    )

    library = _get_document_library(
        client=client,
        site_id=site_id,
        library_name=validated_library_name,
    )

    list_id = library["list_id"]

    api_filters = _build_document_filters(
        site_name=validated_site_name,
        library_name=validated_library_name,
        filters=filters,
        created_before=created_before,
        created_after=created_after,
    )

    url = (
        f"{client.base}"
        f"/sites/{site_id}"
        f"/lists/{list_id}"
        f"/items"
    )

    params = {
        "$expand": "fields,driveItem",
        "$top": validated_page_size,
    }

    if api_filters:
        params["$filter"] = " and ".join(api_filters)

    headers = client.auth.graph_headers()

    # Fortæller Graph, at vi forventer at filtrere på
    # indekserede SharePoint-kolonner.
    headers["Prefer"] = (
        "HonorNonIndexedQueriesWarningMayFailRandomly"
    )

    documents: list[dict[str, Any]] = []
    next_url: str | None = url
    next_params: dict[str, Any] | None = params

    while next_url:
        response = requests.get(
            next_url,
            headers=headers,
            params=next_params,
            timeout=60,
        )
        response.raise_for_status()

        response_data = response.json()

        for item in response_data.get("value", []):
            parsed_document = _parse_document_item(item)

            # Mapper API-felter tilbage til de viste kolonnenavne.
            parsed_document.update(
                _extract_business_fields(
                    site_name=validated_site_name,
                    library_name=validated_library_name,
                    raw_fields=item.get("fields", {}),
                )
            )

            documents.append(parsed_document)

            if (
                validated_max_items is not None
                and len(documents) >= validated_max_items
            ):
                documents = documents[
                    :validated_max_items
                ]
                next_url = None
                break

        if next_url is None:
            break

        next_url = response_data.get("@odata.nextLink")

        # nextLink indeholder allerede query-parametrene.
        next_params = None

    return {
        "exists": bool(documents),
        "count": len(documents),
        "documents": documents,
    }


def delete_document(
    *,
    site_name: str,
    library_name: str,
    drive_item_id: str,
) -> None:
    """
    Sletter ét dokument fra et SharePoint-dokumentbibliotek.

    Parametre:
        site_name:
            Eksempel: "Automatisering".

        library_name:
            Eksempel: "Digitalpost".

        drive_item_id:
            SharePoints tekniske drive item-id.

    Output:
        Funktionen returnerer None ved korrekt sletning.

    Bemærkning:
        SharePoints retention- og papirkurvsregler afgør,
        om filen slettes permanent med det samme.
    """
    validated_site_name = _require_text(
        site_name,
        "site_name",
    )
    validated_library_name = _require_text(
        library_name,
        "library_name",
    )
    validated_drive_item_id = _require_text(
        drive_item_id,
        "drive_item_id",
    )

    client = get_client()

    site_id = client.get_site_id(
        validated_site_name
    )

    library = _get_document_library(
        client=client,
        site_id=site_id,
        library_name=validated_library_name,
    )

    drive_id = library["drive_id"]

    delete_url = (
        f"{client.base}"
        f"/drives/{drive_id}"
        f"/items/{validated_drive_item_id}"
    )

    response = requests.delete(
        delete_url,
        headers=client.auth.graph_headers(),
        timeout=60,
    )

    # 204 er normal succes.
    # 404 behandles også som succes, da dokumentet allerede er væk.
    if response.status_code == 404:
        return

    response.raise_for_status()


def _get_document_library(
    *,
    client,
    site_id: str,
    library_name: str,
) -> dict[str, str]:
    """
    Finder både list-id og drive-id for et dokumentbibliotek.

    Output:
        {
            "list_id": "...",
            "drive_id": "...",
            "name": "Digitalpost",
        }
    """
    lists_url = (
        f"{client.base}"
        f"/sites/{site_id}"
        f"/lists"
    )

    lists_response = requests.get(
        lists_url,
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    lists_response.raise_for_status()

    matching_lists = [
        item
        for item in lists_response.json().get("value", [])
        if item.get("displayName", "").casefold()
        == library_name.casefold()
    ]

    if len(matching_lists) != 1:
        raise RuntimeError(
            f"Forventede præcis ét bibliotek med navnet "
            f"{library_name!r}, men fandt "
            f"{len(matching_lists)}."
        )

    list_item = matching_lists[0]
    list_id = list_item["id"]

    drive_url = (
        f"{client.base}"
        f"/sites/{site_id}"
        f"/lists/{list_id}"
        f"/drive"
    )

    drive_response = requests.get(
        drive_url,
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    drive_response.raise_for_status()

    drive_data = drive_response.json()
    drive_id = drive_data.get("id")

    if not drive_id:
        raise RuntimeError(
            f"Biblioteket {library_name!r} har ikke et drive-id. "
            "Kontrollér, at det er et dokumentbibliotek."
        )

    return {
        "list_id": list_id,
        "drive_id": drive_id,
        "name": list_item.get(
            "displayName",
            library_name,
        ),
    }


def _get_list_item_for_drive_item(
    *,
    client,
    drive_id: str,
    drive_item_id: str,
) -> dict[str, Any]:
    """
    Finder listeelementet, som hører til en fil.
    """
    url = (
        f"{client.base}"
        f"/drives/{drive_id}"
        f"/items/{drive_item_id}"
        f"/listItem"
    )

    response = requests.get(
        url,
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    response.raise_for_status()

    return response.json()


def _get_document_by_drive_item_id(
    *,
    client,
    drive_id: str,
    drive_item_id: str,
) -> dict[str, Any]:
    """
    Henter metadata for én konkret fil.
    """
    url = (
        f"{client.base}"
        f"/drives/{drive_id}"
        f"/items/{drive_item_id}"
    )

    response = requests.get(
        url,
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    response.raise_for_status()

    return response.json()


def _update_document_fields(
    *,
    client,
    site_id: str,
    list_id: str,
    list_item_id: str,
    fields: dict[str, Any],
) -> None:
    """
    Opdaterer kolonnerne på et dokument.
    """
    url = (
        f"{client.base}"
        f"/sites/{site_id}"
        f"/lists/{list_id}"
        f"/items/{list_item_id}"
        f"/fields"
    )

    response = requests.patch(
        url,
        headers=client.auth.graph_headers(),
        json=fields,
        timeout=30,
    )
    response.raise_for_status()


def _map_ui_fields_to_api_fields(
    *,
    site_name: str,
    library_name: str,
    fields: dict[str, Any],
) -> dict[str, Any]:
    """
    Mapper SharePoints viste kolonnenavne til interne API-navne.

    Det er vigtigt, fordi et synligt navn som dokument_id
    godt kan få et internt SharePoint-navn med tegnkodning.
    """
    if not isinstance(fields, dict):
        raise TypeError("fields skal være en dictionary.")

    schema = get_list_schema(
        site_name,
        library_name,
    )

    unknown_fields = [
        field_name
        for field_name in fields
        if field_name not in schema
    ]

    if unknown_fields:
        raise ValueError(
            f"Ukendte SharePoint-felter: {unknown_fields}. "
            f"Gyldige felter: {sorted(schema)}"
        )

    return {
        schema[field_name]["api_name"]: value
        for field_name, value in fields.items()
    }


def _extract_business_fields(
    *,
    site_name: str,
    library_name: str,
    raw_fields: dict[str, Any],
) -> dict[str, Any]:
    """
    Mapper interne API-navne tilbage til viste kolonnenavne.
    """
    schema = get_list_schema(
        site_name,
        library_name,
    )

    result: dict[str, Any] = {}

    for ui_name, metadata in schema.items():
        api_name = metadata["api_name"]
        result[ui_name] = raw_fields.get(api_name)

    return result


def _build_document_filters(
    *,
    site_name: str,
    library_name: str,
    filters: dict[str, Any] | None,
    created_before: datetime | str | None,
    created_after: datetime | str | None,
) -> list:
    """
    Bygger Graph-filtre til dokumentopslag.
    """
    api_filters: list[str] = []

    if filters:
        if not isinstance(filters, dict):
            raise TypeError(
                "filters skal være en dictionary eller None."
            )

        mapped_filters = _map_ui_fields_to_api_fields(
            site_name=site_name,
            library_name=library_name,
            fields=filters,
        )

        for api_name, value in mapped_filters.items():
            escaped_value = _escape_odata_value(value)

            api_filters.append(
                f"fields/{api_name} eq {escaped_value}"
            )

    if created_before is not None:
        created_before_value = _to_utc_iso(
            created_before,
            "created_before",
        )

        api_filters.append(
            f"fields/Created lt '{created_before_value}'"
        )

    if created_after is not None:
        created_after_value = _to_utc_iso(
            created_after,
            "created_after",
        )

        api_filters.append(
            f"fields/Created gt '{created_after_value}'"
        )

    return api_filters


def _parse_document_item(
    item: dict[str, Any],
) -> dict[str, Any]:
    """
    Laver et råt listeelement om til en enkel dictionary.
    """
    drive_item = item.get("driveItem") or {}
    fields = item.get("fields") or {}

    return {
        "list_item_id": item.get("id"),
        "drive_item_id": drive_item.get("id"),
        "file_name": drive_item.get("name"),
        "web_url": drive_item.get("webUrl"),
        "size": drive_item.get("size"),
        "created": (
            drive_item.get("createdDateTime")
            or item.get("createdDateTime")
            or fields.get("Created")
        ),
        "modified": (
            drive_item.get("lastModifiedDateTime")
            or item.get("lastModifiedDateTime")
            or fields.get("Modified")
        ),
        "e_tag": drive_item.get("eTag"),
    }


def _escape_odata_value(value: Any) -> str:
    """
    Gør en Python-værdi sikker til et simpelt OData-filter.
    """
    if value is None:
        return "null"

    if isinstance(value, bool):
        return "true" if value else "false"

    if isinstance(value, (int, float)):
        return str(value)

    escaped_text = str(value).replace(
        "'",
        "''",
    )

    return f"'{escaped_text}'"


def _to_utc_iso(
    value: datetime | str,
    field_name: str,
) -> str:
    """
    Omregner datetime eller ISO-tekst til UTC-tekst.
    """
    if isinstance(value, str):
        cleaned_value = value.strip()

        if cleaned_value.endswith(("Z", "z")):
            cleaned_value = (
                cleaned_value[:-1] + "+00:00"
            )

        try:
            parsed_value = datetime.fromisoformat(
                cleaned_value
            )
        except ValueError as error:
            raise ValueError(
                f"{field_name} har et ugyldigt datoformat."
            ) from error

    elif isinstance(value, datetime):
        parsed_value = value

    else:
        raise TypeError(
            f"{field_name} skal være datetime eller tekst."
        )

    if parsed_value.tzinfo is None:
        raise ValueError(
            f"{field_name} skal indeholde en tidszone."
        )

    utc_value = parsed_value.astimezone(
        timezone.utc
    )

    return (
        utc_value
        .isoformat(timespec="seconds")
        .replace("+00:00", "Z")
    )


def _require_text(
    value: Any,
    field_name: str,
) -> str:
    """
    Kontrollerer et obligatorisk tekstfelt.
    """
    if not isinstance(value, str):
        raise TypeError(
            f"{field_name} skal være tekst."
        )

    validated_value = value.strip()

    if not validated_value:
        raise ValueError(
            f"{field_name} må ikke være tom."
        )

    return validated_value


def _validate_file_name(file_name: str) -> str:
    """
    Kontrollerer at filnavnet er sikkert.
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

    found_invalid_characters = [
        character
        for character in invalid_characters
        if character in validated_file_name
    ]

    if found_invalid_characters:
        raise ValueError(
            f"file_name indeholder ugyldige tegn: "
            f"{sorted(found_invalid_characters)}"
        )

    return validated_file_name


def _validate_content(content: bytes) -> bytes:
    """
    Kontrollerer filindholdet.
    """
    if not isinstance(content, bytes):
        raise TypeError(
            "content skal være bytes."
        )

    if not content:
        raise ValueError(
            "content må ikke være tom."
        )

    return content


def _validate_page_size(page_size: int) -> int:
    """
    Kontrollerer sidestørrelsen.
    """
    if isinstance(page_size, bool):
        raise TypeError(
            "page_size skal være et helt tal."
        )

    if not isinstance(page_size, int):
        raise TypeError(
            "page_size skal være et helt tal."
        )

    if not 1 <= page_size <= MAX_PAGE_SIZE:
        raise ValueError(
            f"page_size skal være mellem 1 og "
            f"{MAX_PAGE_SIZE}."
        )

    return page_size


def _validate_max_items(
    max_items: int | None,
) -> int | None:
    """
    Kontrollerer det valgfrie maksimum.
    """
    if max_items is None:
        return None

    if isinstance(max_items, bool):
        raise TypeError(
            "max_items skal være et helt tal eller None."
        )

    if not isinstance(max_items, int):
        raise TypeError(
            "max_items skal være et helt tal eller None."
        )

    if max_items <= 0:
        raise ValueError(
            "max_items skal være større end 0."
        )

    return max_items