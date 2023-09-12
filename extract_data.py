import datetime
import pathlib
import re
import warnings

warnings.simplefilter(action="ignore", category=FutureWarning)

import pandas as pd

files_dir = pathlib.Path("downloads")

final_df = pd.DataFrame(columns=["FileName", "SiteName", "Cycle", "Date"])

for file in files_dir.iterdir():
    print(file.name)
    corrections = str.maketrans({" ": "", "_": "-", "ٍ": ""})
    site_name = re.search(r"(L|Q|I)(ALX|UPP|CAI)(.)(\d{5})", file.name.upper()).group()
    cycle = re.search(r"C\d{1,2}", file.name.upper()).group()
    date = re.search(
        r"\d{1,2}-\w{3,10}-\d{4}", file.name.translate(corrections)
    ).group()

    try:
        date = datetime.datetime.strptime(date, "%d-%b-%Y")
    except ValueError:
        date = datetime.datetime.strptime(date, "%d-%B-%Y")
    try:
        data = pd.read_excel(file, sheet_name="Serials", skiprows=4)
        cols = [col for col in data.columns if "Unnamed" not in col]
        cols = cols[:4]
        data = data[cols]

        file_data = pd.DataFrame()
        for col_mark in range(2, 6, 2):
            temp_df = data.iloc[:, col_mark - 2 : col_mark].copy()
            temp_df.dropna(inplace=True)
            temp_df.columns = pd.Index(["Item", "Serial"])
            file_data = pd.concat([file_data, temp_df])

        file_data["FileName"] = file.name
        file_data["SiteName"] = site_name
        file_data["Cycle"] = cycle
        file_data["Date"] = date.strftime("%Y-%m-%d")
        file_data["ItemType"] = file_data.Item.apply(
            lambda x: "Battery" if "battery" in x.lower() else "Rectifier"
        )

        final_df = pd.concat([final_df, file_data[file_data.Serial != "#0"]])
    except Exception as e:
        print(f"ERROR: {file.name}: {e}")
final_df["duplicated_flag"] = final_df.duplicated(subset=["Serial"], keep=False)
final_df.Serial = final_df.Serial.str.strip()

final_df.to_excel(
    f"serials_extract_{int(datetime.datetime.now().timestamp())}.xlsx", index=False
)
