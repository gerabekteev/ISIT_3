"""forecasting — прогнозный компонент DemoScope (прототип ЛР3).

Публичный API пакета: реестр алгоритмов и контекст прогнозирования.
"""
from __future__ import annotations

from .base import ForecastStrategy
from .context import ALGORITHMS, ForecastContext, register
from .models import ForecastMetrics, ForecastResult, TimeSeries

__all__ = [
    "ForecastStrategy",
    "ForecastContext",
    "ForecastResult",
    "ForecastMetrics",
    "TimeSeries",
    "ALGORITHMS",
    "register",
]
