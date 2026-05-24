def validate_and_map(data, mapping):
    result = {}

    for key, value in data.items():
        if key not in mapping:
            raise Exception(f"Field '{key}' not found in mapping")

        m = mapping[key]
        sp_name = m["sp_name"]
        t = m["type"]

        if t == "lookup":
            result[sp_name] = {"LookupId": value}
        elif t == "user":
            result[sp_name] = value
        else:
            result[sp_name] = value

    return result


def reverse_map(fields, mapping):
    result = {}
    reverse = {v["sp_name"]: k for k, v in mapping.items()}

    for k, v in fields.items():
        if k in reverse:
            result[reverse[k]] = v

    return result
