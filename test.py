import pandas as pd
import openpyxl
import json
import os

try:
    all_data = []
    for root, dirs, files in os.walk('Data/'):
        for file in files:
            if file.endswith('.json'):
                filepath = os.path.join(root, file)
                province = os.path.basename(root)
                district = file[:-5]  # remove .json
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        if isinstance(data, dict) and "data" in data and "address" in data["data"] and "analysts" in data["data"]["address"]:
                            analysts = data["data"]["address"]["analysts"]
                            for item in analysts:
                                item['province'] = province
                                item['district'] = district
                            all_data.extend(analysts)
                        else:
                            print(f"Unexpected structure in {filepath}")
                except Exception as e:
                    print(f"Error reading {filepath}: {e}")

    if all_data:
        df = pd.DataFrame(all_data)
        columns_to_keep = ['province', 'district', 'month', 'year', 'soTin', 'tyLeTangTruongSoTin', 'giaTrenDtTrungBinh', 'tyLeTangTruongGia']
        df = df[columns_to_keep]
        df.to_excel('data.xlsx', index=False)
        print("Data exported successfully to data.xlsx")
    else:
        print("No data found")

except Exception as e:
    print(f"An error occurred: {e}")