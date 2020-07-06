import openpyxl
from sheet_management import *

# Excel workbook and sheets
workbook = openpyxl.load_workbook("AktienExcel.xlsx")
Status_sheet = workbook.get_sheet_by_name("Status")
Historie_sheet = workbook.get_sheet_by_name("Historie")


def main():

    update_status()
    update_prices()

    print("Willkommen zu Deiner Aktienverwaltung.")

    while True:
        command = menu()

        if command == 0:
            break

        handle_command(command)

        save()

    print("Auf Wiedersehen! Bis zum nächsten Mal.")


# USER INTERFACE

def menu():
    command = -1
    valid_commands = ["1", "2", "3", "0"]

    while command not in valid_commands:

        command = input("\nWas möchtest Du tun? \nAktien hinzufügen(1)  "
                        "Aktien entfernen(2)  Depot anzeigen(3)  Beenden(0)\n")

    return int(command)


def handle_command(command):

    if command == 1 or command == 2:
        update_stock_count_ui(command)
    elif command == 3:
        print_current_stocks()


def update_stock_count_ui(command):
    operations = ["hinzufügen", "entfernen"]

    symbol = input(f"Gib die Wertpapierkennung der Aktie an, die Du {operations[command - 1]} möchtest.\n").upper()
    count = int(input(f"Wie viele {symbol}-Aktien möchtest Du {operations[command - 1]}?\n"))

    if command == 2:
        count = -count

    add_stock(symbol, count)


if __name__ == "__main__":
    main()

