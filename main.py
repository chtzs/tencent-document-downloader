from openpyxl import Workbook
from download import SheetDownloader
from sheet_generator import SheetGenerator


def main():
    url = "https://docs.qq.com/sheet/DVGh0WlNHTmNCUlNH?tab=4hutl0"
    downloader = SheetDownloader(url)
    print("Document name: %s" % downloader.title)
    wb = Workbook()
    for tab in downloader.tabs:
        print("Fetching %s..." % tab["name"])
        content, _, max_col = downloader.fetch_sheet_data(tab["id"])
        _ = SheetGenerator(wb.create_sheet(tab["name"]), content, max_col)

    empty_ws = wb["Sheet"]
    wb.remove(empty_ws)
    wb.save("%s.xlsx" % downloader.title)
    print("Saved to %s.xlsx" % downloader.title)

if __name__ == "__main__":
    main()