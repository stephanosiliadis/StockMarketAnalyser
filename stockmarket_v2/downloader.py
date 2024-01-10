import pandas as pd
import yfinance as yf
import datetime as dt


class DataDownloader(object):
    def __init__(self, tickers):
        self.tickers = tickers
        self.data = None

    def download_data(self):
        year, month, date = map(
            int,
            input(
                "Enter the year, month, date from which to start downloading: "
            ).split(","),
        )
        start = dt.datetime(year, month, date)
        end = dt.datetime.now()
        self.data = yf.download(self.tickers, start, end)
        return self.data

    def get_xlsx(self, filename):
        pd.DataFrame(self.data).to_excel(filename)

