from openpyxl import Workbook
from download import SheetDownloader
from sheet_generator import SheetGenerator

from sys import argv

USAGE = """
USAGE: python3 main.py <tencent document url> [cookie file]
"""

def try_get_url():
    if len(argv) == 1:
        print(USAGE)
        exit(1)
    else:
        url = argv[1]
        return url

def main():
    url = try_get_url()
    downloader = SheetDownloader(url)
    print("Document name: %s" % downloader.title)
    wb = Workbook()
    for tab in downloader.tabs:
        print("Fetching sheet %s..." % tab["name"])
        content, _, max_col = downloader.fetch_sheet_data(tab["id"])
        sheet = SheetGenerator(wb.create_sheet(tab["name"]), content, max_col)
        sheet.generate_sheet()

    print("Generating...")

    empty_ws = wb["Sheet"]
    wb.remove(empty_ws)
    wb.save("%s.xlsx" % downloader.title)
    print("Saved to %s.xlsx" % downloader.title)


if __name__ == "__main__":
    main()
