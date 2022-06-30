import msb as msb
import pandas as pd
from datetime import datetime
from util import format_excel_date, format_excel_float


def main():
    wb = msb.Workbook("templates/excel")
    wb_new = wb.copy()
    today = format_excel_date(datetime.now())

    df = pd.DataFrame({"date": ["1/1/2022", "2/1/2022"], "value": [0.58, 0.45]})
    df["date"] = pd.to_datetime(df["date"]).apply(format_excel_date)
    df["value"] = df["value"].apply(format_excel_float)

    ws = msb.Worksheet(1)
    ws.stream_and_dump_template(df=df, today=today)

    wb_new.stream_and_dump_template(df=df)
    wb_new.to_xlsx("Book1")


if __name__ == "__main__":
    main()
