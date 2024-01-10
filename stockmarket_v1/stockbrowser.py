import requests
from requests.exceptions import ConnectionError
from bs4 import BeautifulSoup
from tqdm import tqdm


class StockBrowser(object):
    def __init__(self, stocks, headers):
        self.stock_data = {}
        self.stocks_history = {}
        self.latest_history = {}
        self.stocks = stocks
        self.headers = headers

    def web_data_div(self, web_data, class_path):
        web_data_div = web_data.find_all(
            "div",
            {"class": class_path},
        )
        try:
            spans = web_data_div[0].find_all("span")
            texts = [span.get_text() for span in spans]

        except IndexError:
            texts = None

        return texts

    def web_data_span(self, web_data, data_bind):
        web_data_span = web_data.find(
            "span",
            {"data-bind": data_bind},
        )

        return web_data_span.get_text()

    def browse_stock(self, stock):
        url = f"https://www.capital.gr/finance/historycloses/{stock}/"
        try:
            response = requests.get(url)
            web_data = BeautifulSoup(response.text, "lxml")
            texts = self.web_data_div(web_data, "finance__details__left")
            if texts != None:
                price, performance, percent_performance = (
                    texts[0].replace(",", "."),
                    texts[1].replace(",", "."),
                    texts[2].replace(")", "").replace("(", "").replace(",", "."),
                )

            else:
                price, performance, percent_performance = None, None, None

            open_price = self.web_data_span(
                web_data,
                "text: o.extend({priceNumeric: {dependOn: l, forceSign: true, precision: 4 }})",
            )

            high = self.web_data_span(
                web_data,
                "text: hp.extend({priceNumeric: {dependOn: l, forceSign: true, precision: 4 }})",
            )

            low = self.web_data_span(
                web_data,
                "text: lp.extend({priceNumeric: {dependOn: l, forceSign: true, precision: 4 }})",
            )

            volume = self.web_data_span(
                web_data,
                "text: tv.extend({numeric: { precision: 0}}), flashBackground: tv",
            )

            value = (
                self.web_data_span(
                    web_data,
                    "text: to.extend({numeric: { precision: 0 }})() + ' €', visible: to() > 0, flashBackground: to",
                )
                .strip()
                .replace("\r\n", "")
            )

            actions = self.web_data_span(
                web_data,
                "text: t.extend({numeric: { precision: 0}}), flashBackground: t",
            )

        except ConnectionError:
            (
                price,
                performance,
                percent_performance,
                open_price,
                high,
                low,
                volume,
                value,
                actions,
            ) = (
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
            )

        return (
            price,
            performance,
            percent_performance,
            open_price,
            high,
            low,
            volume,
            value,
            actions,
        )

    def browse_stock_history(self, stock):
        url = f"https://www.capital.gr/finance/historycloses/{stock}/"
        stock_history = {}
        try:
            response = requests.get(url)
            web_data = BeautifulSoup(response.text, "lxml")
            data_div = web_data.find(
                "div",
                {"class": "cp__table"},
            )
            data_tbody = data_div.find("tbody")
            table_rows = data_tbody.find_all("tr")
            for tr in table_rows:
                table_data = [td.get_text() for td in tr.find_all("td")]
                date = table_data[0]
                stock_history[date] = {
                    "Close": table_data[1],
                    "% Performance": table_data[2],
                    "Market Open": table_data[3],
                    "High": table_data[4],
                    "Low": table_data[5],
                    "Volume": table_data[6],
                    "Value": table_data[7],
                }

        except ConnectionError:
            pass

        return stock_history

    def browse_stock_latest_history(self, stock):
        url = f"https://www.capital.gr/finance/historycloses/{stock}/"
        stock_history = {}
        try:
            response = requests.get(url)
            web_data = BeautifulSoup(response.text, "lxml")
            data_div = web_data.find(
                "div",
                {"class": "cp__table"},
            )
            data_tbody = data_div.find("tbody")
            tr = data_tbody.find("tr")
            td = [td.get_text() for td in tr.find_all("td")]
            stock_history[td[0]] = {
                "Close": td[1],
                "% Performance": td[2],
                "Market Open": td[3],
                "High": td[4],
                "Low": td[5],
                "Volume": td[6],
                "Value": td[7],
            }

        except ConnectionError:
            pass

        return stock_history

    def get_stock_data(self):
        print("Getting stock data...")
        for stock in tqdm(self.stocks):
            self.stock_data[stock] = self.headers.copy()
            data = self.browse_stock(stock)
            for idx, header in enumerate(self.stock_data[stock].keys()):
                self.stock_data[stock][header] = data[idx]

        return self.stock_data

    def get_stock_history(self):
        print("Getting stock history...")
        for stock in tqdm(self.stocks):
            self.stocks_history[stock] = self.browse_stock_history(stock)

        return self.stocks_history

    def get_stock_latest_history(self):
        print("Getting stock latest history...")
        for stock in tqdm(self.stocks):
            self.latest_history[stock] = self.browse_stock_latest_history(stock)

        return self.latest_history
