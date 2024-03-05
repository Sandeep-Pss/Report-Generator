import json
import pandas as pd


def flatten_json(json_data, parent_key='', sep='_'):
    items = {}
    for k, v in json_data.items():
        new_key = parent_key + sep + k if parent_key else k
        if isinstance(v, dict):
            flattened = flatten_json(v, new_key, sep)
            items.update(flattened)
        elif isinstance(v, list):
            for i, item in enumerate(v, start=1):
                if isinstance(item, dict):
                    flattened = flatten_json(item, new_key, sep)
                    for key, value in flattened.items():
                        items[key] = items.get(key, []) + [value]
                else:
                    items[new_key] = ', '.join(str(elem) for elem in v)
        else:
            items[new_key] = v
    return items


def json_to_excel(json_file, key=None):
    # Read JSON file
    with open(json_file, 'r') as f:
        data = json.load(f)

    # If a key is specified, access the corresponding data
    if key:
        data = data.get(key, [])

    # If data is a list, iterate over each object and flatten it
    if isinstance(data, list):
        flattened_data = [flatten_json(obj) for obj in data]
    else:
        flattened_data = flatten_json(data)

    # Convert to DataFrame
    df = pd.DataFrame(flattened_data)

    # Write to Excel
    excel_file = json_file.replace('.json', '.xlsx')
    df.to_excel(excel_file, index=False)


