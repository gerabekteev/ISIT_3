"""cohort_component.py — когортно-компонентная стратегия прогнозирования.

Классический демографический метод: население продвигается по годам
поколений (когортам) с учётом рождаемости, смертности и миграции.
Прототип реализует однолетний шаг модели средствами стандартной
библиотеки; возрастная структура и коэффициенты задаются данными
или синтезируются из итоговой численности (стационарное приближение).
"""
from __future__ import annotations

import math

from .base import ForecastStrategy
from .metrics import compute_metrics
from .models import ForecastResult, TimeSeries

Pyramid = dict[str, dict[int, float]]

FERTILE_AGES = range(15, 50)   # фертильный возраст женщин
TFR = 1.5                      # суммарный коэффициент рождаемости
SEX_RATIO_MALE = 0.512         # доля мальчиков среди новорождённых


class CohortComponentStrategy(ForecastStrategy):
    """Прогноз когортно-компонентным методом (рождаемость/смертность/миграция)."""

    name = "cohort"

    def __init__(
        self,
        age_structure: Pyramid | None = None,
        net_migration: float | None = None,
        tfr: float = TFR,
    ) -> None:
        self.age_structure = age_structure
        self.net_migration = net_migration
        self.tfr = tfr

    def forecast(self, series: TimeSeries, horizon: int = 10) -> ForecastResult:
        """Строит прогноз когортно-компонентным методом на horizon лет.

        В отличие от статистических стратегий доверительные интервалы
        не формируются (результат детерминирован при заданных
        коэффициентах); качество оценивается бэктестом по всей истории.
        """
        horizon = max(5, min(horizon, 15))
        # Пирамида строится по первому значению ряда; бэктест проходит
        # всю историю, прогноз продолжает траекторию модели от её конца
        pyramid = self.age_structure or self._synthesize_pyramid(series.values[0])

        # Калибровка миграции: фактическое среднегодовое изменение ряда
        # минус естественный прирост модели (рождаемость минус смертность)
        migration = (
            self.net_migration
            if self.net_migration is not None
            else self._calibrate_migration(series, pyramid)
        )

        # Честная проверка качества: проходим весь исторический период
        # и сравниваем траекторию модели с фактической численностью
        history, state_end = self._project(pyramid, len(series.values) - 1, migration)
        metrics = compute_metrics(list(series.values[1:]), history)

        forecast_totals, _ = self._project(state_end, horizon, migration)
        years = tuple(range(series.years[-1] + 1, series.years[-1] + 1 + horizon))

        return ForecastResult(
            model_name="Когортно-компонентный метод",
            forecast_years=years,
            forecast_values=tuple(round(v, 1) for v in forecast_totals),
            metrics=metrics,
            params={"шаг модели": "1 год", "СКР": self.tfr,
                    "миграция, тыс./год": round(migration, 2)},
        )

    # ── Компоненты модели ──

    def _project(self, pyramid: Pyramid, years: int, net_migration: float) -> tuple[list[float], Pyramid]:
        """Прогон модели на years однолетних шагов.

        Returns:
            (итоги численности на каждом шаге, состояние пирамиды на конец периода).
        """
        totals: list[float] = []
        for _ in range(years):
            births = self._births(pyramid)
            pyramid = self._advance(pyramid)
            self._apply_migration(pyramid, net_migration)
            pyramid["male"][0] = pyramid["male"].get(0, 0.0) + births * SEX_RATIO_MALE
            pyramid["female"][0] = pyramid["female"].get(0, 0.0) + births * (1.0 - SEX_RATIO_MALE)
            totals.append(sum(sum(ages.values()) for ages in pyramid.values()))
        return totals, pyramid

    def _births(self, pyramid: Pyramid) -> float:
        """Число рождений за год: СКР, распределённый по фертильным возрастам женщин."""
        women = sum(pyramid["female"].get(age, 0.0) for age in FERTILE_AGES)
        return women * self.tfr / len(FERTILE_AGES)

    @staticmethod
    def _advance(pyramid: Pyramid) -> Pyramid:
        """Сдвиг когорт на год вперёд с учётом дожития (модель Гомперца).

        Упрощение прототипа: возраст когорт ограничен 100 годами —
        выжившие в возрасте 100+ накапливаются в закрытой группе «100+».
        """
        advanced: Pyramid = {"male": {}, "female": {}}
        for sex in ("male", "female"):
            for age, count in pyramid[sex].items():
                survivors = count * CohortComponentStrategy._survival(age, sex)
                if survivors <= 0:
                    continue
                next_age = min(age + 1, 100)
                advanced[sex][next_age] = advanced[sex].get(next_age, 0.0) + survivors
        return advanced

    @staticmethod
    def _survival(age: int, sex: str) -> float:
        """Вероятность дожития (интенсивность смертности — модель Гомперца-Мейкема)."""
        base = 0.00018 * math.exp(0.0705 * age)
        base += 0.005 if age == 0 else 0.0            # младенческая смертность
        if sex == "male":
            base *= 1.5                                # сверхсмертность мужчин
        return max(1.0 - min(base, 0.5), 0.0)

    @staticmethod
    def _apply_migration(pyramid: Pyramid, net_migration: float) -> None:
        """Нетто-миграция: равномерно по трудоспособным возрастам обоих полов."""
        if abs(net_migration) < 1e-9:
            return
        working = list(range(18, 60))
        per_age = net_migration / (2 * len(working))
        for sex in ("male", "female"):
            for age in working:
                pyramid[sex][age] = max(pyramid[sex].get(age, 0.0) + per_age, 0.0)

    def _calibrate_migration(self, series: TimeSeries, pyramid: Pyramid) -> float:
        """Калибровка миграции: средний фактический прирост минус естественный."""
        deltas = [b - a for a, b in zip(series.values, series.values[1:])]
        mean_delta = sum(deltas) / len(deltas)
        natural_first_year, _ = self._project(pyramid, 1, 0.0)
        natural = natural_first_year[0] - sum(sum(a.values()) for a in pyramid.values())
        return mean_delta - natural

    @staticmethod
    def _synthesize_pyramid(total: float) -> Pyramid:
        """Стационарное приближение пирамиды по итоговой численности.

        Возрастные веса воспроизводят типичную структуру населения
        депопулирующего региона РФ (доля 65+ — около 20 %).
        """
        def weight(age: int, sex: str) -> float:
            if age < 15:
                base = 0.92
            elif age < 55:
                base = 1.0
            elif age < 65:
                base = 0.82
            elif age < 75:
                base = 0.62 if sex == "male" else 0.76
            elif age < 85:
                base = 0.28 if sex == "male" else 0.45
            else:
                base = 0.08 if sex == "male" else 0.16
            return base

        weights = {
            sex: [weight(age, sex) for age in range(101)]
            for sex in ("male", "female")
        }
        # Мужчин в популяции меньше (перевес женщин в старших возрастах)
        male_sum = sum(weights["male"])
        female_sum = sum(weights["female"])
        male_share = male_sum / (male_sum + female_sum)
        pyramid: Pyramid = {}
        for sex, w in weights.items():
            sex_sum = male_sum if sex == "male" else female_sum
            scale = total * (male_share if sex == "male" else 1 - male_share) / sex_sum
            pyramid[sex] = {age: w[age] * scale for age in range(101)}
        return pyramid
