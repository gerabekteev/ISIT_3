"""holt_winters.py — стратегия прогнозирования Хольта-Уинтерса.

Двойное экспоненциальное сглаживание (модель Хольта) с сеточным
подбором параметров по информационному критерию AIC. Прототип
воспроизводит логику боевого модуля forecasting/holt_winters.py
проекта (statsmodels.ExponentialSmoothing) средствами стандартной
библиотеки: перебор типов тренда (add/mul) и демпфирования тренда.
"""
from __future__ import annotations

import math
from itertools import product

from .base import ForecastStrategy
from .metrics import compute_metrics
from .models import ForecastResult, TimeSeries

_ALPHAS = (0.2, 0.4, 0.6, 0.8)
_BETAS = (0.1, 0.3, 0.5)
_PHIS = (0.85, 0.95, 1.0)

Params = dict[str, str | float | bool]


class HoltWintersStrategy(ForecastStrategy):
    """Прогноз методом двойного экспоненциального сглаживания."""

    name = "holt_winters"

    def forecast(self, series: TimeSeries, horizon: int = 10) -> ForecastResult:
        """Строит прогноз двойным экспоненциальным сглаживанием с ДИ 95%."""
        horizon = max(5, min(horizon, 15))
        y: list[float] = list(series.values)
        if len(y) < 4:
            raise ValueError(f"Недостаточно данных: {len(y)} точек (нужно минимум 4)")

        fitted, level, trend, params, aic = self._select_model(y)

        # Прогноз вперёд: уровень фиксируется, тренд затухает с коэффициентом phi
        phi = float(params["phi"])
        multiplicative = params["trend"] == "mul"
        forecast_values = [
            max(
                level * (1.0 + sum(phi ** k for k in range(1, h + 1)) * trend)
                if multiplicative
                else level + sum(phi ** k for k in range(1, h + 1)) * trend,
                0.0,
            )
            for h in range(1, horizon + 1)
        ]

        metrics = compute_metrics(y[1:], fitted[1:])
        resid_std = math.sqrt(
            sum((a - p) ** 2 for a, p in zip(y[1:], fitted[1:])) / max(len(y) - 1, 1)
        )
        ci = [1.96 * resid_std * math.sqrt(h) for h in range(1, horizon + 1)]

        return ForecastResult(
            model_name="Хольт-Уинтерс (двойное экспоненциальное сглаживание)",
            forecast_years=tuple(range(series.years[-1] + 1, series.years[-1] + 1 + horizon)),
            forecast_values=tuple(round(v, 1) for v in forecast_values),
            lower_ci=tuple(round(v - c, 1) for v, c in zip(forecast_values, ci)),
            upper_ci=tuple(round(v + c, 1) for v, c in zip(forecast_values, ci)),
            metrics=metrics,
            params={"trend": str(params["trend"]), "damped": bool(params["damped"]),
                    "alpha": float(params["alpha"]), "beta": float(params["beta"]),
                    "phi": phi, "aic": round(aic, 1)},
        )

    # ── Подбор и оценка модели ──

    def _select_model(
        self, y: list[float]
    ) -> tuple[list[float], float, float, Params, float]:
        """Сеточный перебор параметров; выбор конфигурации с минимумом AIC."""
        best: tuple[list[float], float, float, Params, float] | None = None
        for trend_type, damped, alpha, beta in product(
            ("add", "mul"), (True, False), _ALPHAS, _BETAS
        ):
            if trend_type == "mul" and any(v <= 0 for v in y):
                continue
            phis = _PHIS if damped else (1.0,)
            for phi in phis:
                params: Params = {
                    "trend": trend_type, "damped": damped,
                    "alpha": alpha, "beta": beta, "phi": phi,
                }
                fitted, level, trend, sse = self._fit(y, params)
                if sse is None or sse <= 0:
                    continue
                k = 3 + (1 if damped else 0)  # alpha, beta, sigma2 (+ phi при демпфировании)
                aic = len(y) * math.log(sse / len(y)) + 2 * k
                if best is None or aic < best[4]:
                    best = (fitted, level, trend, params, aic)
        if best is None:
            params = {"trend": "add", "damped": True, "alpha": 0.6,
                      "beta": 0.3, "phi": 0.95}
            fitted, level, trend, _ = self._fit(y, params)
            best = (fitted, level, trend, params, float("inf"))
        return best

    def _fit(
        self, y: list[float], params: Params
    ) -> tuple[list[float], float, float, float | None]:
        """Проход рекуррентных уравнений Хольта по ряду.

        Returns:
            Кортеж (сглаженные значения, конечный уровень, конечный тренд, SSE).
        """
        alpha, beta, phi = float(params["alpha"]), float(params["beta"]), float(params["phi"])
        multiplicative = params["trend"] == "mul"
        level = y[0]
        trend = (y[1] - y[0]) / y[0] if (multiplicative and y[0] != 0) else (y[1] - y[0])
        fitted: list[float] = [level]
        sse = 0.0
        for t in range(1, len(y)):
            if multiplicative:
                forecast_point = level * (1.0 + phi * trend)
            else:
                forecast_point = level + phi * trend
            error = y[t] - forecast_point
            sse += error ** 2
            if multiplicative and level != 0:
                new_level = alpha * y[t] + (1 - alpha) * level * (1.0 + phi * trend)
                new_trend = beta * (new_level / level - 1.0) + (1 - beta) * phi * trend
            else:
                new_level = alpha * y[t] + (1 - alpha) * (level + phi * trend)
                new_trend = beta * (new_level - level) + (1 - beta) * phi * trend
            level, trend = new_level, new_trend
            fitted.append(forecast_point)
        return fitted, level, trend, sse
