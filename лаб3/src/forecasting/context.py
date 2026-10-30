"""context.py — контекст прогнозирования (паттерн «Стратегия»).

Хранит текущую стратегию, обеспечивает её динамическую подмену
во время выполнения и единый метод построения прогноза.
Клиентский код (маршрутизатор бэкенда) не знает, какой алгоритм
используется, — он оперирует только контрактом ForecastStrategy.
"""
from __future__ import annotations

from .base import ForecastStrategy
from .models import ForecastResult, TimeSeries

# Реестр доступных алгоритмов: имя -> класс стратегии
ALGORITHMS: dict[str, type[ForecastStrategy]] = {}


def register(cls: type[ForecastStrategy]) -> type[ForecastStrategy]:
    """Декоратор регистрации стратегии в реестре алгоритмов."""
    ALGORITHMS[cls.name] = cls
    return cls


class ForecastContext:
    """Контекст: делегирует расчёт прогноза активной стратегии."""

    def __init__(self, strategy: ForecastStrategy) -> None:
        self._strategy = strategy

    @property
    def strategy_name(self) -> str:
        """Имя активного алгоритма (для логирования и UI)."""
        return self._strategy.name

    def set_strategy(self, strategy: ForecastStrategy) -> None:
        """Динамическая подмена алгоритма во время выполнения."""
        self._strategy = strategy

    def run_forecast(self, series: TimeSeries, horizon: int = 10) -> ForecastResult:
        """Единая точка входа: расчёт прогноза активной стратегией."""
        return self._strategy.forecast(series, horizon)
