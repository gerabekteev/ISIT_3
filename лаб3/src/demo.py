"""demo.py — демонстрация работы прогнозного компонента.

Сценарий: один и тот же исторический ряд численности населения
прогоняется через единый ForecastContext; стратегии берутся из
реестра ALGORITHMS — добавление нового алгоритма в реестр не требует
изменения клиентского кода (демонстрация Open/Closed паттерна
«Стратегия»).

Запуск: python demo.py
"""
from __future__ import annotations

from forecasting.arima import ArimaStrategy
from forecasting.cohort_component import CohortComponentStrategy
from forecasting.context import ALGORITHMS, ForecastContext, register
from forecasting.holt_winters import HoltWintersStrategy
from forecasting.models import TimeSeries

register(HoltWintersStrategy)
register(ArimaStrategy)
register(CohortComponentStrategy)

# Демонстрационный ряд: численность населения депопулирующего региона,
# тыс. чел. (структура данных воспроизводит ряды проекта DemoScope)
DEMO_YEARS = tuple(range(2010, 2026))
DEMO_VALUES = (1553.9, 1544.5, 1542.1, 1539.7, 1536.0, 1530.1, 1523.2,
               1514.8, 1505.3, 1494.6, 1483.4, 1470.9, 1461.4, 1453.2,
               1445.1, 1437.0)


def main() -> None:
    series = TimeSeries(DEMO_YEARS, DEMO_VALUES)
    context = ForecastContext(HoltWintersStrategy())

    print("=" * 78)
    print("DemoScope: прогноз численности населения региона на 5 лет")
    print(f"Исторический ряд: {series.years[0]}-{series.years[-1]} "
          f"({series.values[0]:.1f} -> {series.values[-1]:.1f} тыс. чел.)")
    print("=" * 78)

    # Клиентский код знает только реестр и единый контекст —
    # конкретные стратегии подставляются во время выполнения
    for strategy_cls in ALGORITHMS.values():
        context.set_strategy(strategy_cls())
        result = context.run_forecast(series, horizon=5)

        print(f"\nАлгоритм: {result.model_name} [{context.strategy_name}]")
        if result.params:
            print(f"  Параметры: {result.params}")
        if result.metrics:
            print(f"  Метрики аппроксимации: MAPE={result.metrics.mape}% "
                  f"RMSE={result.metrics.rmse} MAE={result.metrics.mae}")
        print("  Прогноз (тыс. чел.):")
        rows = list(zip(result.forecast_years, result.forecast_values,
                        result.lower_ci or [None] * len(result.forecast_values),
                        result.upper_ci or [None] * len(result.forecast_values)))
        for year, value, lo, hi in rows:
            if lo is None or hi is None:
                print(f"    {year}: {value:8.1f}")
            else:
                print(f"    {year}: {value:8.1f}   (95% ДИ: {lo:8.1f} .. {hi:8.1f})")

    print("\n" + "=" * 78)
    print(f"Проверка SOLID: контекст принял {len(ALGORITHMS)} стратегий(и) из реестра")
    print("ALGORITHMS без изменения клиентского кода — принцип Open/Closed выполнен.")
    print("=" * 78)


if __name__ == "__main__":
    main()
