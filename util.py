import numpy as np
import pandas as pd
from datetime import date


def format_excel_date(dt: date) -> int:
    "Convert a date to an Excel date."
    # That 1900 was not a leap year is noted in pandas but not excel,
    # so one extra day needs to be added.
    return (pd.to_datetime(dt) - pd.Timestamp("1900-01-01")).days + 2


def format_excel_float(x: float) -> str:
    "Format a floating-point scalar as a decimal string in positional or scientific notation."
    if x >= 1:
        return np.format_float_positional(x, min_digits=16)
    if x >= 0.1:
        return np.format_float_positional(x, min_digits=17)
    return np.format_float_scientific(x, exp_digits=1, min_digits=16).replace("e", "E")
