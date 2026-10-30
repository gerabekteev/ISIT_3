"""models.py — модели данных прогнозного компонента DemoScope.

Контракты данных, которыми обмениваются стратегии прогнозирования
и контекст. Все структуры неизменяемы (frozen) — данные результата
создаются один раз и не модифицируются потребителями.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class TimeSeries:
    """Временной ряд показателя: годы и значения одинаковой длины."""

    years: tuple[int, ...]
    values: tuple[float, ...]

    def __post_init__(self) -> None:
        if len(self.years) != len(self.values):
            raise ValueError("Длины years и values должны совпадать")
        if any(b <= a for a, b in zip(self.years, self.years[1:])):
            raise ValueError("Годы должны строго возрастать")


@dataclass(frozen=True)
class ForecastMetrics:
    """Метрики качества аппроксимации на исторических данных."""

    mape: float  # средняя абсолютная процентная ошибка, %
    rmse: float  # среднеквадратическая ошибка
    mae: float   # средняя абсолютная ошибка


@dataclass(frozen=True)
class ForecastResult:
    """Единый контракт результата прогноза для всех стратегий."""

    model_name: str
    forecast_years: tuple[int, ...]
    forecast_values: tuple[float, ...]
    lower_ci: tuple[float, ...] = field(default=())
    upper_ci: tuple[float, ...] = field(default=())
    metrics: ForecastMetrics | None = None
    params: dict[str, str | int | float | bool] = field(default_factory=dict)
