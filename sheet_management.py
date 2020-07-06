from main import *
from stock_information import *
import collections
from datetime import date

# MANAGING STOCKS

def save():
    workbook.save("AktienExcel.xlsx")


def get_stock_count(symbol):

    for cell in Status_sheet["A"]:

        if cell.value == symbol:
            return Status_sheet.cell(cell.row, cell.column + 1).value

    return -1


def get_current_stocks():

    global stocks
    global stocks_have_been_updated

    if stocks is None or stocks_have_been_updated:
        set_current_stocks()

    return stocks


def set_current_stocks():

    current_stocks = []
    Stock = collections.namedtuple("Stock", "name count price")

    for cell in Status_sheet["A"][1:]:

        if cell.value is not None:

            stock_append = Stock(cell.value, Status_sheet.cell(cell.row, cell.column + 1).value, get_price(cell.value))
            current_stocks.append(stock_append)

    global stocks
    stocks = current_stocks


def print_current_stocks():

    print("Das aktuelle Depot wird geladen. Dies kann eine Weile dauern.")

    stocks = get_current_stocks()

    print("Aktien:")

    for stock in stocks:

        padded_name = " " * (4 - len(stock.name)) + stock.name
        padded_count = " " * (2 - len(str(stock.count))) + str(stock.count)
        print(f"{padded_name}: {padded_count} Stück à {stock.price}€")


def stock_is_valid(symbol):

    try:
        stock = Stock(symbol, token=IEX_TOKEN)
        stock.get_company_name().__str__()
        return True
    except Exception:
        print(f"Werpapierkennung '{symbol}' konnte nicht gefunden werden.")
        return False


def add_stock(symbol, count):

    symbol = symbol.upper()

    if stock_is_valid(symbol):
        stock_cell_row = find_cell(symbol)
        # 2 because stock count is in second column of status sheet
        amount_cell = Status_sheet.cell(stock_cell_row, 2)
        update_stock_count(Status_sheet.cell(stock_cell_row, 1), amount_cell, symbol, count)


def find_cell(symbol):
    """Support method for add_stock. Finds either the first empty cell row or the cell row corresponding to the stock
    name. Very specific with hardcoded values, only for this single purpose."""

    row = 0

    for cell in Status_sheet["A"]:

        row += 1

        if cell.value == symbol:
            return row

    row += 1

    return row


def update_stock_count(stock_cell, amount_cell, symbol, count):
    """Support method for add_stock. Updates the stock count of stocks to be added or removed."""

    global stocks_have_been_updated

    if stock_cell.value == symbol:

        if amount_cell.value + count <= 0:
            print(f"Du besitzt nun keine {symbol}-Aktien mehr.")
            amount_cell.value = amount_cell.value + count
            cleanup_stocks()
            stocks_have_been_updated = True
            return

        amount_cell.value = amount_cell.value + count

    else:

        if count > 0:
            stock_cell.value = symbol
            amount_cell.value = count
        else:
            print("Die Anzahl der hinzuzufügenden Aktien muss positiv sein.")
            return

    stocks_have_been_updated = True
    print(f"Die neue Anzahl an {symbol}-Aktien beträgt {get_stock_count(symbol)}.")


def cleanup_stocks():
    """Finds all the stocks that have a count of 0 or less and removes them from the sheet."""

    range_to_move = find_range_to_move()
    if range_to_move[0] < len(Status_sheet["B"]):
        move_string = f"A{range_to_move[0]}:B{range_to_move[1]}"
        Status_sheet.move_range(move_string, rows=-1, cols=0)

    # this condition is true if the deleted stock is in the last row
    elif range_to_move[0] > range_to_move[1]:
        Status_sheet.delete_rows(range_to_move[0])


def find_range_to_move():
    """Support method for cleanup_stocks. Finds the range of cells to be moved after stock has been deleted."""

    start = find_next_empty_cell(2) + 1
    end = find_next_empty_cell(start) - 1

    return [start, end]


def find_next_empty_cell(start):
    """Support method for find_range_to_move. Very specific with hardcoded values."""

    index = start - 1

    for row in range(start, len(Status_sheet["B"]) + 2):

        value = Status_sheet.cell(row, 2).value

        if value is None or value <= 0:
            return index

        index += 1


def find_next_empty_cell_vertical(sheet, column, start):
    """Support method. Not in use right now, just for consistency.
    Finds the next empty cell after the cell with coordinates (start, column)."""

    index = start - 1
    index_at_end = True

    for cell in sheet.iter_rows(max_col=column, min_row=start, values_only=True):
        index += 1

        if cell[0] is None:
            index_at_end = False
            break

    if index_at_end:
        index += 1

    return index


def find_next_empty_cell_horizontal(sheet, row, start):
    """Support method. Finds the next empty cell after the cell with coordinates (row, start)."""

    index = start - 1
    index_at_end = True

    for cell in sheet.iter_cols(max_row=row, min_col=start, values_only=True):
        index += 1

        if cell[0] is None:
            index_at_end = False
            break

    if index_at_end:
        index += 1

    return index


def get_performances_euro():

    performances = []
    base_prices = get_base_prices()
    current_stocks = get_current_stocks()

    for index, base_price in enumerate(base_prices):
        stock = current_stocks[index]
        performance = stock.price * stock.count - base_price
        performances.append(float(str("%.2f" % performance)))

    return performances


def get_performances_percent():
    pass


def get_base_prices():

    base_values = []

    for cell in Status_sheet.iter_rows(min_row=2, min_col=5, max_col=5, values_only=True):
        if cell[0] is None:
            break

        base_values.append(cell[0])

    return base_values


# UPDATING STOCKS

def update_workbook():

    update_prices()
    update_status()


def update_prices():

    today = get_formatted_date()
    stocks = get_current_stocks()
    column = find_next_empty_cell_horizontal(Historie_sheet, 1, 1)

    row = 1
    Historie_sheet.cell(row, column).value = today

    for stock in stocks:
        row += 1
        enter_formatted_in_euro(Historie_sheet.cell(row, column), stock.price)

    save()


def update_status():

    performances = get_performances_euro()
    stocks = get_current_stocks()

    for index, performance in enumerate(performances):
        enter_formatted_in_euro(Status_sheet.cell(index + 2, 6), stocks[index].price)
        enter_formatted_in_euro(Status_sheet.cell(index + 2, 7), performance)

    save()


def enter_formatted_in_euro(cell, amount):
    cell.number_format = "#,##0.00€"
    cell.value = amount


def get_formatted_date():

    today = date.today()
    return f"{today.day}.{today.month}.{today.year}"

