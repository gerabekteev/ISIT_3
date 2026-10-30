"""base.py — абстрактный контракт стратегии прогнозирования.

Паттерн «Стратегия»: единый интерфейс, скрывающий конкретный
алгоритм расчёта прогноза за методом forecast(). Новые алгоритмы
добавляются наследованием без изменения клиентского кода (OCP).
"""
from __future__ import annotations

from abc import ABC, abstractmethod

from .models import ForecastResult, TimeSeries


class ForecastStrategy(ABC):
    """Интерфейс стратегии прогнозирования динамики населения."""

    name: str = "strategy"

    @abstractmethod
    def forecast(self, series: TimeSeries, horizon: int = 10) -> ForecastResult:
        """Строит прогноз на horizon лет по историческому ряду.

        Args:
            series: исторический временной ряд показателя.
            horizon: горизонт прогнозирования в годах (5..15).

        Returns:
            ForecastResult: прогнозные значения, доверительные
            интервалы и метрики аппроксимации.
        """
        raise NotImplementedError
