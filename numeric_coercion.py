from __future__ import annotations

from typing import Union

import pandas as pd


def _strip_if_str(value):
    return value.strip() if isinstance(value, str) else value


def normalize_numeric_like(obj: Union[pd.Series, pd.DataFrame]) -> Union[pd.Series, pd.DataFrame]:
    """Normalize common human-formatted numeric strings.

    - Strips surrounding whitespace
    - Treats empty/whitespace-only strings as missing (pd.NA)
    - Removes thousand separators (",")
    - Removes percent signs ("%")

    Notes
    -----
    - Percent signs are *removed* but values are NOT divided by 100.
      E.g. "12%" -> "12" -> 12.0
    """

    if isinstance(obj, pd.DataFrame):
        cleaned = obj.apply(lambda col: col.map(_strip_if_str))
        cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True)
        cleaned = cleaned.replace({",": "", "%": ""}, regex=True)
        return cleaned

    if isinstance(obj, pd.Series):
        cleaned = obj.map(_strip_if_str)
        cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True)
        cleaned = cleaned.replace({",": "", "%": ""}, regex=True)
        return cleaned

    raise TypeError(f"Unsupported type: {type(obj)!r}")


def coerce_numeric_series(series: pd.Series, *, errors: str = "coerce") -> pd.Series:
    """Coerce a series to numeric after normalizing numeric-like strings."""

    cleaned = normalize_numeric_like(series)
    return pd.to_numeric(cleaned, errors=errors)
