def to_camel(name):
    parts = name.split()
    return parts[0].lower() + "".join(p.capitalize() for p in parts[1:])


def normalize(col):
    if col.get("choice"):
        return "choice"
    if col.get("personOrGroup"):
        return "user"
    if col.get("lookup"):
        return "lookup"
    if col.get("dateTime"):
        return "date"
    return "text"


def generate_mapping(columns):
    mapping = {}

    for col in columns:
        if col.get("readOnly"):
            continue

        key = to_camel(col["displayName"])

        mapping[key] = {
            "sp_name": col["name"],
            "type": normalize(col)
        }

    return mapping
