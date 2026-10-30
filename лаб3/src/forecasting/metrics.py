"""metrics.py — метрики качества прогнозных моделей.

Чистые функции без состояния; используются всеми стратегиями
для расчёта ошибок аппроксимации на историческом участке ряда.
"""
from __future__ import annotations

import math

from .models import ForecastMetrics


def mape(actual: list[float], predicted: list[float]) -> float:
    """Средняя абсолютная процентная ошибка (%); точки с нулём исключаются."""
    pairs = [(a, p) for a, p in zip(actual, predicted) if a != 0]
    if not pairs:
        return 0.0
    return sum(abs(a - p) / abs(a) for a, p in pairs) / len(pairs) * 100.0


def rmse(actual: list[float], predicted: list[float]) -> float:
    """Среднеквадратическая ошибка."""
    return math.sqrt(sum((a - p) ** 2 for a, p in zip(actual, predicted)) / len(actual))


def mae(actual: list[float], predicted: list[float]) -> float:
    """Средняя абсолютная ошибка."""
    return sum(abs(a - p) for a, p in zip(actual, predicted)) / len(actual)


def compute_metrics(actual: list[float], predicted: list[float]) -> ForecastMetrics:
    """Сводный расчёт метрик качества модели."""
    return ForecastMetrics(
        mape=round(mape(actual, predicted), 2),
        rmse=round(rmse(actual, predicted), 2),
        mae=round(mae(actual, predicted), 2),
    )
