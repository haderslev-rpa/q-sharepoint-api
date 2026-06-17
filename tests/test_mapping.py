from src.q_sharepoint_api.sp_mapping import validate_and_map

mapping = {
    "customerName": {"sp_name": "Title", "type": "text"},
    "status": {"sp_name": "Status", "type": "choice"}
}

data = {
    "customerName": "Test",
    "status": "Open"
}

result = validate_and_map(data, mapping)

print(result)