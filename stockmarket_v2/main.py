from downloader import DataDownloader
from handler import ExcelHandler


def main():
    tickers = [
        "KARE.AT",
        "ANDRO.AT",
        "INKAT.AT",
    ]
    filename = "Stocks.xlsx"

    downloader = DataDownloader(tickers)
    handler = ExcelHandler(filename, downloader)
    handler.style_cells()
    handler.align_cells()
    handler.round_values()
    handler.save_changes()


if __name__ == "__main__":
    main()
