"""Sletning af ét SharePoint-listeitem."""

import requests

from q_sharepoint_api.sp_api import get_client


def delete_list_item(
    site_name: str,
    list_name: str,
    item_id: str | int,
) -> None:
    """Sletter ét SharePoint-listeitem.

    Funktionen returnerer None, når SharePoint bekræfter sletningen.
    """

    if not isinstance(site_name, str) or not site_name.strip():
        raise ValueError(
            "site_name skal være en ikke-tom tekst."
        )

    if not isinstance(list_name, str) or not list_name.strip():
        raise ValueError(
            "list_name skal være en ikke-tom tekst."
        )

    validated_item_id = str(item_id).strip()

    if not validated_item_id.isdigit():
        raise ValueError(
            "item_id skal være et positivt helt tal."
        )

    client = get_client()

    site_id = client.get_site_id(
        site_name.strip()
    )

    list_id = client.get_list_id(
        site_id,
        list_name.strip(),
    )

    url = (
        f"{client.base}"
        f"/sites/{site_id}"
        f"/lists/{list_id}"
        f"/items/{validated_item_id}"
    )

    response = requests.delete(
        url,
        headers=client.auth.graph_headers(),
        timeout=30,
    )

    if response.status_code != 204:
        response.raise_for_status()