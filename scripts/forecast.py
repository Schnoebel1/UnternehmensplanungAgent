from __future__ import annotations
import numpy as np
from typing import List

def cagr(start: float, end: float, years: int) -> float:
    if years <= 0 or start in (None, 0, np.nan) or end in (None, np.nan):
        return 0.0
    try:
        return (end / start) ** (1 / years) - 1
    except ZeroDivisionError:
        return 0.0

def project(end_val: float, growth: float, horizon: int = 3) -> List[float]:
    return [end_val * (1 + growth) ** i for i in range(1, horizon + 1)]
