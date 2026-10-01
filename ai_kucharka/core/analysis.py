import os
from collections.abc import Sequence
from itertools import zip_longest
from typing import Any

import numpy as np
import pandas as pd

from .data import InfoStr


def apply_conversion(xs: Sequence[InfoStr], conversion: dict[str, Any], check_lower: bool = False) -> list[InfoStr]:

    output = []
    for x in xs:
        x_test = x.value.lower() if check_lower else x.value
        if x_test in conversion.keys():
            output.append(InfoStr(conversion[x_test], x.source))
        elif x_test in conversion.values():
            output.append(InfoStr(x_test, x.source))
    return output


def convert_list(x):
    if len(x) == 0:
        return np.nan
    elif len(x) == 1:
        return x[0]
    else:
        return x


def expand_row(row: pd.Series, cols: list[str]) -> pd.DataFrame:
    max_len = max(len(row[col]) if isinstance(row[col], list) else 1 for col in cols)
    if max_len <= 1:
        return pd.DataFrame(row).T

    expanded = []
    for values in zip_longest(*(row[col] for col in cols), fillvalue=None):
        new_row = row.copy()
        for col, val in zip(cols, values):
            if val is None:
                new_row[col] = []
            else:
                new_row[col] = [val]
        expanded.append(new_row)
    return pd.DataFrame(expanded)


def expand_rows(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    df = pd.concat(expand_row(row, cols) for _, row in df.iterrows())
    return df.reset_index(drop=True)


def find_images(root: str, img_extensions: tuple[str, ...] = (".png", ".jpg", ".jpeg")) -> pd.DataFrame:

    data = []
    for path, directories, files in os.walk(root):
        for file in files:
            if file.lower().endswith(tuple(img_extensions)):
                data.append({"path": os.path.relpath(path, start=root), "file": file})
    return pd.DataFrame(data)


def remove_nonunique(infos: list[InfoStr]) -> list[InfoStr]:
    unique = {}
    for info in infos:
        if info.value not in unique:
            unique[info.value] = info
    return list(unique.values())


def remove_strings(infos: list[InfoStr], ignore_str: set[str] = set()) -> list[InfoStr]:
    unique = {}
    for info in infos:
        x_lower = info.value.lower()
        if x_lower not in unique and x_lower not in ignore_str:
            unique[x_lower] = info

    return [info for x_lower, info in unique.items() if not any(x_lower != y_lower and x_lower.startswith(y_lower) for y_lower in unique.keys())]
