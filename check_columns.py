
import glob
import json
import pandas as pd

column_names = ["Order No", "Customer Nr", "Material", "CD Comment"]
names_dict = {}

for file in glob.glob(r"C:\Temp\output\*.csv"):
    df = pd.read_csv(file, error_bad_lines=False)
    all_data = pd.DataFrame()
    all_data.append(df, ignore_index=True)
    columns_file = list(df.columns.values)
    columns_file = columns_file[0].split(";")
    names_dict[file]=[]
    for name in columns_file:
        for pattern in column_names:
            if name.lower() == pattern.lower():
                names_dict[file].append(name)

json_file = json.dumps(names_dict)
with open(r"C:\Temp\output\differences.json", "w") as new_file:
    new_file.write(json_file)
