import os
import re
import pathlib

from dotenv import load_dotenv
from O365 import Account

load_dotenv()


def authenticate() -> Account:
    credentials = (
        os.getenv("CLIENT_ID"),
        os.getenv("CLIENT_SECRET"),
    )

    scopes = ["Sites.Read.All", "offline_access", "User.Read"]
    account = Account(credentials, scopes=scopes)
    if not account.is_authenticated:
        print("Please login to your Office 365 account")
        account.authenticate()
    return account


account = authenticate()
download_path = pathlib.Path("downloads")
if not download_path.exists():
    download_path.mkdir()

site = account.sharepoint().get_site(os.getenv("SITE_ID"))
doc_lib = site.get_default_document_library()
years = [2023]
for year in years:
    path = f"/PM_Trackers/Packages/{year}"
    year_folder = doc_lib.get_item_by_path(path)
    week_folders = year_folder.get_child_folders()
    for week_folder in week_folders:
        packages_folders = week_folder.get_child_folders()
        for package_folder in packages_folders:
            try:
                site_package = package_folder.get_items()
                for item in site_package:
                    match_excel = re.search(r"\.xls[x]?$", item.name)
                    if item.is_file and match_excel:
                        item.download(to_path=download_path, name=f"{year}-{item.name}")
                        print(item.name)
            except Exception as e:
                print(e)
print("finished")
