"""arima.py — стратегия прогнозирования ARIMA.

Прототип семейства ARIMA(p, d, 0) с сеточным подбором параметров
по критерию AIC: ряд дифференцируется d раз, авторегрессия порядка
p оценивается методом наименьших квадратов (нормальные уравнения
решаются методом Гаусса). MA-составляющая в прототипе опущена;
боевой модуль проекта (statsmodels.ARIMA) перебирает полную сетку
(p, d, q) — логика выбора лучшей модели воспроизведена.
"""
from __future__ import annotations

import math

from .base import ForecastStrategy
from .metrics import compute_metrics
from .models import ForecastResult, TimeSeries


class ArimaStrategy(ForecastStrategy):
    """Прогноз моделью ARIMA(p, d, 0) с подбором порядка по AIC."""

    name = "arima"

    def forecast(self, series: TimeSeries, horizon: int = 10) -> ForecastResult:
        """Строит прогноз ARIMA(p,d,0) с подбором параметров по AIC и ДИ 95%."""
        horizon = max(5, min(horizon, 15))
        y: list[float] = list(series.values)
        if len(y) < 5:
            raise ValueError(f"Недостаточно данных для ARIMA: {len(y)} точек (нужно минимум 5)")

        # Порядок интегрирования d: ряд с выраженным линейным трендом
        # (|r| Пирсона > 0.85) берётся в разностях, иначе — в уровнях;
        # порядок p подбирается сеточным поиском по AIC при фиксированном d
        d = 1 if abs(self._trend_correlation(y)) > 0.85 else 0
        diff = self._difference(y, d)
        best: tuple[int, int, list[float], float, float] | None = None
        # (p, d, diff, coeffs, aic)
        for p in range(0, 4):
            if len(diff) <= p + 2:
                continue
            coeffs, residuals, sse = self._fit_ar(diff, p)
            if coeffs is None or residuals is None or sse is None or sse <= 0:
                continue
            aic = len(diff) * math.log(sse / len(diff)) + 2 * (p + 1)
            if best is None or aic < best[4]:
                best = (p, d, diff, coeffs, aic)
        if best is None:  # страхующая конфигурация ARIMA(1, 1, 0)
            diff = self._difference(y, 1)
            coeffs = self._fit_ar(diff, 1)[0]
            if coeffs is None:
                coeffs = [0.0, 0.0]
            best = (1, 1, diff, coeffs, float("inf"))
        p, d, diff, coeffs, aic = best

        # Итеративный прогноз дифференцированного ряда по AR(p), затем интегрирование
        window: list[float] = list(diff[-p:]) if p > 0 else []
        diff_forecast: list[float] = []
        for _ in range(horizon):
            nxt = coeffs[-1] + sum(c * v for c, v in zip(coeffs[:-1], window[-p:] if p > 0 else []))
            diff_forecast.append(nxt)
            window.append(nxt)
        forecast_values = self._integrate(y[-1], diff_forecast, d)

        # Аппроксимация на историческом участке: для t >= p разности
        # предсказываются моделью, затем восстанавливаются в уровни
        pdiff = [
            coeffs[-1] + sum(c * v for c, v in zip(coeffs[:-1], diff[t - p:t]))
            for t in range(p, len(diff))
        ]
        residuals = [diff[t] - pdiff[i] for i, t in enumerate(range(p, len(diff)))]
        resid_std = math.sqrt(sum(r ** 2 for r in residuals) / max(len(residuals), 1))
        ci = [1.96 * resid_std * math.sqrt(h) for h in range(1, horizon + 1)]

        if d == 0:
            fitted_tail = pdiff
            historical_tail = y[p:]
        else:
            fitted_tail, acc = [], y[p]
            for v in pdiff:
                acc += v
                fitted_tail.append(acc)
            historical_tail = y[p + 1:]
        metrics = compute_metrics(historical_tail, fitted_tail) if fitted_tail else None

        return ForecastResult(
            model_name=f"ARIMA({p},{d},0)",
            forecast_years=tuple(range(series.years[-1] + 1, series.years[-1] + 1 + horizon)),
            forecast_values=tuple(round(v, 1) for v in forecast_values),
            lower_ci=tuple(round(v - c, 1) for v, c in zip(forecast_values, ci)),
            upper_ci=tuple(round(v + c, 1) for v, c in zip(forecast_values, ci)),
            metrics=metrics,
            params={"p": p, "d": d, "q": 0, "aic": round(aic, 1)},
        )

    # ── Вспомогательные операции ──

    @staticmethod
    def _trend_correlation(series: list[float]) -> float:
        """Коэффициент корреляции Пирсона между номером точки и значением."""
        n = len(series)
        xs = list(range(n))
        mx, my = sum(xs) / n, sum(series) / n
        cov = sum((x - mx) * (v - my) for x, v in zip(xs, series))
        sx = math.sqrt(sum((x - mx) ** 2 for x in xs))
        sy = math.sqrt(sum((v - my) ** 2 for v in series))
        return cov / (sx * sy) if sx > 0 and sy > 0 else 0.0

    @staticmethod
    def _difference(series: list[float], d: int) -> list[float]:
        """d-кратное дифференцирование ряда."""
        out = series
        for _ in range(d):
            out = [b - a for a, b in zip(out, out[1:])]
        return out

    @staticmethod
    def _integrate(last_value: float, diff_values: list[float], d: int) -> list[float]:
        """Обратное интегрирование разностей в уровни ряда.

        d=0 — значения возвращаются как есть; d=1 — кумулятивная сумма
        от опорного значения. Прототип ограничен d <= 1 (порядок d выше
        не выбирается детектором тренда).
        """
        if d >= 2:
            raise ValueError("Прототип ARIMA поддерживает порядок интегрирования d <= 1")
        if d <= 0:
            return list(diff_values)
        out: list[float] = []
        acc = last_value
        for v in diff_values:
            acc += v
            out.append(acc)
        return out

    @staticmethod
    def _fit_ar(
        series: list[float], p: int
    ) -> tuple[list[float] | None, list[float] | None, float | None]:
        """Оценка AR(p) с константой методом наименьших квадратов.

        Returns:
            (коэффициенты [c_1..c_p, intercept], остатки, SSE)
            либо (None, None, None) при вырожденной системе.
        """
        n = len(series) - p
        if n <= p + 2:
            return None, None, None
        x = [[series[t + j] for j in range(p)] + [1.0] for t in range(n)]
        target = [series[t + p] for t in range(n)]
        k = p + 1
        normal = [[sum(x[i][a] * x[i][b] for i in range(n)) for b in range(k)] for a in range(k)]
        rhs = [sum(x[i][a] * target[i] for i in range(n)) for a in range(k)]
        coeffs = ArimaStrategy._solve_gauss(normal, rhs)
        if coeffs is None:
            return None, None, None
        residuals = [target[i] - sum(coeffs[j] * x[i][j] for j in range(k)) for i in range(n)]
        sse = sum(r ** 2 for r in residuals)
        return coeffs, residuals, sse

    @staticmethod
    def _solve_gauss(a: list[list[float]], b: list[float]) -> list[float] | None:
        """Решение СЛАУ методом Гаусса с частичным выбором ведущего элемента."""
        n = len(b)
        m = [row[:] + [b[i]] for i, row in enumerate(a)]
        for col in range(n):
            pivot = max(range(col, n), key=lambda r: abs(m[r][col]))
            if abs(m[pivot][col]) < 1e-12:
                return None
            m[col], m[pivot] = m[pivot], m[col]
            for r in range(col + 1, n):
                factor = m[r][col] / m[col][col]
                for c in range(col, n + 1):
                    m[r][c] -= factor * m[col][c]
        result = [0.0] * n
        for r in range(n - 1, -1, -1):
            result[r] = (m[r][n] - sum(m[r][c] * result[c] for c in range(r + 1, n))) / m[r][r]
        return result
