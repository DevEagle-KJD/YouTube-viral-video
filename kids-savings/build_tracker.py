"""Build Kids Savings.numbers (native Apple Numbers) with full transaction history.

numbers-parser cannot write formulas, so balance cells are left for the
5 formulas described in README.md.
"""
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from numbers_parser import RGB, Alignment, Document, NegativeNumberStyle

AVELYN = [
    ("7/1/2026", "15.00"), ("7/15/2026", "20.00"), ("7/15/2026", "5.00"),
    ("8/5/2026", "15.00"), ("8/5/2026", "2.50"), ("8/8/2026", "-5.00"),
    ("8/16/2026", "10.00"), ("9/1/2026", "15.00"), ("9/3/2026", "-2.25"),
    ("9/9/2026", "-4.00"), ("9/17/2026", "-2.50"), ("9/24/2026", "-1.50"),
    ("10/1/2026", "15.00"), ("10/3/2026", "10.00"), ("10/8/2026", "5.00"),
    ("10/14/2026", "-16.00"),
]
SAMUEL = [
    ("7/1/2026", "5.00", ""), ("7/24/2026", "1.00", ""), ("7/27/2026", "5.00", "Tooth"),
    ("8/5/2026", "5.00", ""), ("8/5/2026", "3.50", ""), ("8/9/2026", "1.00", ""),
    ("8/16/2026", "10.00", ""), ("8/27/2026", "10.00", "Teeth"),
    ("8/28/2026", "16.00", "Grades"), ("9/1/2026", "5.00", ""),
    ("9/3/2026", "-2.25", ""), ("9/10/2026", "-4.50", ""), ("9/16/2026", "-4.50", ""),
    ("9/17/2026", "-4.50", ""), ("9/24/2026", "-1.50", ""), ("10/1/2026", "5.00", ""),
]

# Independent verification of the stated totals
a_total = sum(Decimal(a) for _, a in AVELYN)
s_total = sum(Decimal(a) for _, a, _ in SAMUEL)
assert len(AVELYN) + len(SAMUEL) == 32
assert a_total == Decimal("81.25"), a_total
assert s_total == Decimal("49.25"), s_total
assert a_total + s_total == Decimal("130.50")



def main():
    GREEN = RGB(30, 130, 60)
    RED = RGB(200, 40, 40)
    NAVY = RGB(30, 50, 90)

    doc = Document(sheet_name="Overview", table_name="Overview", num_rows=4, num_cols=2)

    header = doc.add_style(name="Ledger Header", bold=True, font_size=13.0,
                           font_color=RGB(255, 255, 255), bg_color=NAVY,
                           alignment=Alignment("center", "middle"))
    big_label = doc.add_style(name="Big Label", bold=True, font_size=18.0, font_color=NAVY,
                              alignment=Alignment("left", "middle"))
    big_value = doc.add_style(name="Big Value", bold=True, font_size=24.0, font_color=GREEN,
                              alignment=Alignment("right", "middle"))
    deposit = doc.add_style(name="Deposit", font_color=GREEN, font_size=12.0,
                            alignment=Alignment("right", "middle"))
    withdrawal = doc.add_style(name="Withdrawal", font_color=RED, font_size=12.0,
                               alignment=Alignment("right", "middle"))
    body = doc.add_style(name="Ledger Text", font_size=12.0, alignment=Alignment("left", "middle"))
    todo = doc.add_style(name="Formula Needed", italic=True, font_size=11.0,
                         font_color=RGB(150, 150, 150), alignment=Alignment("right", "middle"))


    def money(table, row, col):
        table.set_cell_formatting(row, col, "currency", currency_code="USD", decimal_places=2,
                                  negative_style=NegativeNumberStyle.MINUS,
                                  show_thousands_separator=True)


    def header_row(table, labels):
        for c, label in enumerate(labels):
            table.write(0, c, label, style=header)
        table.row_height(0, 30)


    # ---- Overview -------------------------------------------------------------
    ov = doc.sheets["Overview"].tables["Overview"]
    ov.num_header_cols = 0
    header_row(ov, ["Savings", "Balance"])
    for r, name in enumerate(["Avelyn", "Samuel", "Combined Total"], start=1):
        ov.write(r, 0, name, style=big_label)
        ov.write(r, 1, "formula needed", style=todo)
        ov.row_height(r, 44)
    ov.col_width(0, 200)
    ov.col_width(1, 160)


    # ---- Child sheets ---------------------------------------------------------
    def child_sheet(name, columns, rows):
        doc.add_sheet(name, f"{name} Savings", num_rows=1, num_cols=2)
        sheet = doc.sheets[name]
        summary = sheet.tables[f"{name} Savings"]
        summary.num_header_rows = 0
        summary.num_header_cols = 0
        summary.write(0, 0, "Current Savings", style=big_label)
        summary.write(0, 1, "formula needed", style=todo)
        summary.row_height(0, 50)
        summary.col_width(0, 200)
        summary.col_width(1, 160)

        t = sheet.add_table(f"{name} Transactions", x=0, y=110,
                            num_rows=len(rows) + 1, num_cols=len(columns),
                            num_header_rows=1, num_header_cols=0)
        header_row(t, columns)
        amount_col = columns.index("Amount")
        for r, row in enumerate(rows, start=1):
            date, amount = row[0], Decimal(row[1])
            t.write(r, 0, datetime.strptime(date, "%m/%d/%Y"), style=body)
            t.set_cell_formatting(r, 0, "datetime", date_time_format="M/d/yyyy")
            t.write(r, amount_col, float(amount), style=deposit if amount >= 0 else withdrawal)
            money(t, r, amount_col)
            if "Note" in columns and row[2]:
                t.write(r, columns.index("Note"), row[2], style=body)
            t.row_height(r, 26)
        t.col_width(0, 110)
        for c in range(1, len(columns)):
            t.col_width(c, 140 if columns[c] == "Note" else 110)


    child_sheet("Avelyn", ["Date", "Amount", "Balance"], AVELYN)
    child_sheet("Samuel", ["Date", "Amount", "Note", "Balance"], SAMUEL)

    out = Path(__file__).with_name("Kids Savings.numbers")
    doc.save(out)
    print(f"Saved {out}  Avelyn={a_total} Samuel={s_total} Combined={a_total + s_total}")


if __name__ == "__main__":
    main()
