import csv
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference


DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
CATEGORIES = ("food", "games", "other")
GRAPH_WIDTH = 20
LEGACY_DATA_FILE = Path(__file__).with_name("pocket_money.csv")
WORKBOOK_FILE = Path(__file__).with_name("pocket_money.xlsx")
DATA_HEADERS = (
    "recorded_at",
    "day",
    "money_received",
    *CATEGORIES,
    "daily_balance",
    "savings_goal",
)


def get_amount(prompt):
    while True:
        try:
            amount = float(input(prompt))
            if amount < 0:
                print("Please enter an amount that is zero or more.")
                continue
            return amount
        except ValueError:
            print("Please enter a valid number.")


def track_week():
    daily_records = []

    for day in DAYS:
        print(f"\n{day}")
        money_received = get_amount("Pocket money received: ")
        daily_spent = 0
        category_spending = {}

        for category in CATEGORIES:
            category_spent = get_amount(f"Spent on {category}: ")
            category_spending[category] = category_spent
            daily_spent += category_spent

        daily_balance = money_received - daily_spent
        daily_records.append(
            {
                "day": day,
                "money_received": money_received,
                **category_spending,
                "daily_balance": daily_balance,
            }
        )
        print(f"Balance for {day}: {daily_balance:.2f}")

    return daily_records


def print_summary(daily_records, savings_goal):
    weekly_received = sum(record["money_received"] for record in daily_records)
    weekly_spent = sum(
        record[category] for record in daily_records for category in CATEGORIES
    )
    category_totals = {
        category: sum(record[category] for record in daily_records)
        for category in CATEGORIES
    }
    remaining_balance = weekly_received - weekly_spent
    amount_to_goal = savings_goal - remaining_balance

    print("\nWeekly summary")
    print(f"Pocket money received: {weekly_received:.2f}")
    print(f"Money spent: {weekly_spent:.2f}")
    for category in CATEGORIES:
        print(f"Spent on {category}: {category_totals[category]:.2f}")
    print(f"Remaining balance: {remaining_balance:.2f}")

    if amount_to_goal > 0:
        print(f"You need {amount_to_goal:.2f} more to reach your savings goal.")
    else:
        print("You reached your savings goal!")


def print_daily_savings_graph(daily_records):
    largest_balance = max(
        (abs(record["daily_balance"]) for record in daily_records),
        default=0,
    )

    print("\nDaily savings graph (# = relative amount)")
    for record in daily_records:
        balance = record["daily_balance"]
        bar_length = 0
        if balance and largest_balance:
            bar_length = max(
                1,
                round(abs(balance) / largest_balance * GRAPH_WIDTH),
            )
        if balance > 0:
            bar = f"{'#' * bar_length:>{GRAPH_WIDTH}}|"
        elif balance < 0:
            bar = f"|{'#' * bar_length:<{GRAPH_WIDTH}}"
        else:
            bar = f"{' ' * GRAPH_WIDTH}|"
        print(f"{record['day']:<9} {balance:>8.2f} {bar}")


def append_data_rows(data_sheet, records):
    for record in records:
        data_sheet.append([record[header] for header in DATA_HEADERS])


def load_legacy_csv(data_sheet):
    with LEGACY_DATA_FILE.open(newline="", encoding="utf-8") as data_file:
        for record in csv.DictReader(data_file):
            data_sheet.append(
                [
                    record["recorded_at"],
                    record["day"],
                    float(record["money_received"]),
                    *(float(record[category]) for category in CATEGORIES),
                    float(record["daily_balance"]),
                    float(record["savings_goal"]),
                ]
            )


def save_week(daily_records, savings_goal):
    recorded_at = datetime.now().astimezone().isoformat()
    records = [
        {
            "recorded_at": recorded_at,
            **record,
            "savings_goal": savings_goal,
        }
        for record in daily_records
    ]

    if WORKBOOK_FILE.exists():
        workbook = load_workbook(WORKBOOK_FILE)
        if "Data" in workbook.sheetnames:
            data_sheet = workbook["Data"]
        else:
            data_sheet = workbook.create_sheet("Data", 0)
            data_sheet.append(DATA_HEADERS)
    else:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Data"
        data_sheet.append(DATA_HEADERS)
        if LEGACY_DATA_FILE.exists():
            load_legacy_csv(data_sheet)

    if data_sheet.max_row == 1 and data_sheet["A1"].value is None:
        for column, header in enumerate(DATA_HEADERS, start=1):
            data_sheet.cell(row=1, column=column, value=header)
    elif tuple(
        data_sheet.cell(row=1, column=column).value
        for column in range(1, len(DATA_HEADERS) + 1)
    ) != DATA_HEADERS:
        raise ValueError(f"{WORKBOOK_FILE.name} has an unexpected Data sheet format.")

    append_data_rows(data_sheet, records)

    if "Graph" in workbook.sheetnames:
        workbook.remove(workbook["Graph"])
    graph_sheet = workbook.create_sheet("Graph")
    graph_sheet["A1"] = "Daily savings for the latest week"
    graph_sheet.append(("Day", "Daily savings"))
    for record in daily_records:
        graph_sheet.append((record["day"], record["daily_balance"]))

    chart = LineChart()
    chart.title = "Daily Savings Over the Week"
    chart.x_axis.title = "Day"
    chart.y_axis.title = "Savings"
    chart.add_data(
        Reference(graph_sheet, min_col=2, min_row=2, max_row=2 + len(daily_records)),
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(graph_sheet, min_col=1, min_row=3, max_row=2 + len(daily_records))
    )
    chart.style = 13
    chart.marker = "circle"
    graph_sheet.add_chart(chart, "D2")

    workbook.save(WORKBOOK_FILE)
    print(f"Weekly data and graph saved to {WORKBOOK_FILE.name}.")


def main():
    daily_records = track_week()
    savings_goal = get_amount("\nWhat is your savings goal for the week? ")
    print_summary(daily_records, savings_goal)
    print_daily_savings_graph(daily_records)
    save_week(daily_records, savings_goal)


if __name__ == "__main__":
    main()
