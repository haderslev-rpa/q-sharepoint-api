"""
ATS BOX WRAPPER – SHAREPOINT LIST ITEMS
"""

from q_sharepoint_api.functionality.sp_list_item import (
    get_sharepoint_list_item,
    create_update_sharepoint_list_item
)


# -------------------------------------------------
# 📥 GET → ATS BOX
# -------------------------------------------------
def get_to_ats_box(
    site_name,
    list_name,
    list_item_id,
    *,
    item
):

    sharepoint_data = get_sharepoint_list_item(
        site_name=site_name,
        list_name=list_name,
        list_item_id=list_item_id
    )

    _set_box(item, sharepoint_data)

    return sharepoint_data


# -------------------------------------------------
# 💾 CREATE / UPDATE → FRA ATS BOX
# -------------------------------------------------
def create_or_update_from_ats_box(
    site_name,
    list_name,
    *,
    item,
    robot_kommentar: str | None = None
):

    box = item.data.get("box", {})
    sharepoint_data = box.get("sharepoint", {})

    if not isinstance(sharepoint_data, dict):
        raise Exception("box.sharepoint mangler eller er ikke en dict")

    payload = dict(sharepoint_data)

    # ✅ VIGTIGT: SEND VIA _robot_comment (IKKE direkte felt)
    if robot_kommentar:
        payload["_robot_comment"] = robot_kommentar

    result = create_update_sharepoint_list_item(
        site_name=site_name,
        list_name=list_name,
        sharepoint_data=payload
    )

    _set_box(item, result)

    return result


# -------------------------------------------------
# 🆕 CREATE
# -------------------------------------------------
def create_from_ats_box(
    site_name,
    list_name,
    *,
    item,
    robot_kommentar: str | None = None
):

    box = item.data.get("box", {})
    sharepoint_data = box.get("sharepoint", {})

    if not isinstance(sharepoint_data, dict):
        raise Exception("box.sharepoint mangler eller er ikke en dict")

    payload = dict(sharepoint_data)

    payload.pop("id", None)

    if robot_kommentar:
        payload["_robot_comment"] = robot_kommentar

    result = create_update_sharepoint_list_item(
        site_name=site_name,
        list_name=list_name,
        sharepoint_data=payload
    )

    _set_box(item, result)

    return result


# -------------------------------------------------
# 🔄 UPDATE
# -------------------------------------------------
def update_from_ats_box(
    site_name,
    list_name,
    *,
    item,
    robot_kommentar: str | None = None
):

    box = item.data.get("box", {})
    sharepoint_data = box.get("sharepoint", {})

    if "id" not in sharepoint_data:
        raise Exception("id mangler")

    payload = dict(sharepoint_data)

    if robot_kommentar:
        payload["_robot_comment"] = robot_kommentar

    result = create_update_sharepoint_list_item(
        site_name=site_name,
        list_name=list_name,
        sharepoint_data=payload
    )

    _set_box(item, result)

    return result


# -------------------------------------------------
# HELPER
# -------------------------------------------------
def _set_box(item, sharepoint_data):

    if "box" not in item.data or not isinstance(item.data["box"], dict):
        item.data["box"] = {}

    item.data["box"]["sharepoint"] = sharepoint_data
