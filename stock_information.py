from iexfinance.stocks import Stock
from forex_python.converter import CurrencyRates


# GENERAL DATA
# private auth key to access stock data
# IEX_TOKEN =

# caching variables that are given values after one access to the data
usd_to_eur = None
stocks = None
stocks_have_been_updated = False


def get_price(symbol):
    """Returns the price of a stock given the symbol."""

    stock_string = get_stock_string(symbol)
    price_in_dollar = float(extract_price_string(stock_string))
    price_in_euro = dollars_to_euros(price_in_dollar)

    return float("%.2f" % price_in_euro)


def get_stock_string(symbol):
    """Returns a string representation of the stock given the symbol."""

    stock = Stock(symbol, token=IEX_TOKEN)
    stock_string = stock.get_quote().__str__()

    return stock_string


def extract_price_string(stock_data):
    """Support method for get_price. Extracts the price string from the stock (quote) data. Hardcoded values."""

    start = "'latestPrice': "
    end = ","
    stock_data = stock_data[stock_data.find(start) + len(start):]
    stock_data = stock_data[:stock_data.find(end)]

    return stock_data


def dollars_to_euros(dollars):

    global usd_to_eur

    # because the API to get the rates is really slow, on initial startup the conversion rate is fetched
    # and then gets cached so the next conversion will be seamless
    if usd_to_eur is None:
        set_usd_to_eur_rate()

    return dollars * usd_to_eur


def set_usd_to_eur_rate():
    """Support method. Caches the conversion rate in a global variable."""

    converter = CurrencyRates()
    global usd_to_eur
    usd_to_eur = converter.convert("USD", "EUR", 1)

