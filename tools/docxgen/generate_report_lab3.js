// Генератор отчёта по ЛР3 (docx-js) — клон generate_report.js + листинги кода
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб3";
const OUT = path.join(LAB, "отчет_ЛР3.docx");
const LAB_NUM = "3";

// ── Оформление (по примеру) ──
const FONT = "Times New Roman";
const CODE_FONT = "Courier New";
const BODY_SIZE = 24;   // 12pt
const TABLE_SIZE = 24;  // 12pt
const CODE_SIZE = 18;   // 9pt
const H1_SIZE = 28;     // 14pt
const H2_SIZE = 26;     // 13pt
const LINE_BODY = 360;  // 1,5 интервала
const INDENT = 709;     // 1,25 см красная строка
const BLACK = "000000";
const PAGE = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } };

const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };
const UNDERLINE = { top: NB, left: NB, right: NB, bottom: { style: BorderStyle.SINGLE, size: 4, color: BLACK } };

const run = (text, opts = {}) => new TextRun({ text, size: BODY_SIZE, color: BLACK, font: FONT, ...opts });
const para = (children, opts = {}) => new Paragraph({ children, ...opts });
const empty = (n = 1) => Array.from({ length: n }, () => para([]));

function labelValueRow(cells, widths) {
  let wIdx = 0;
  return new TableRow({
    cantSplit: true,
    children: cells.map(c => {
      const span = c.span || 1;
      const w = (widths || []).slice(wIdx, wIdx + span).reduce((a, b) => a + b, 0);
      wIdx += span;
      return new TableCell({
        columnSpan: span,
        width: { size: w, type: WidthType.PERCENTAGE },
        borders: c.underline ? UNDERLINE : noBorders,
        margins: { top: 20, bottom: 20, left: 40, right: 40 },
        children: [
          para([run(c.t || "", { size: c.size || BODY_SIZE, bold: c.bold })],
            { alignment: c.align || AlignmentType.LEFT, spacing: { line: 276 } }),
          ...(c.hint ? [para([run(c.hint, { size: 16 })], { alignment: AlignmentType.CENTER, spacing: { line: 240 } })] : []),
        ],
      });
    }),
  });
}
function tumanTable(rows, widths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows,
  });
}

function buildTitlePage() {
  const children = [];
  children.push(para([]));
  [
    "МИНОБРНАУКИ РОССИИ",
    "федеральное государственное автономное образовательное учреждение",
    "высшего образования",
    "«Московский государственный технологический университет «СТАНКИН»",
    "(ФГАОУ ВО «МГТУ «СТАНКИН»)",
  ].forEach(t => children.push(para([run(t, { bold: true })], { alignment: AlignmentType.CENTER, spacing: { line: 276 } })));

  children.push(...empty(1));
  children.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE }, borders: noBorders,
          margins: { top: 20, bottom: 20, left: 40, right: 40 },
          children: [para([run("Институт информационных технологий", { size: 22 })], { alignment: AlignmentType.LEFT, spacing: { line: 276 } })],
        }),
        new TableCell({
          width: { size: 50, type: WidthType.PERCENTAGE }, borders: noBorders,
          margins: { top: 20, bottom: 20, left: 40, right: 40 },
          children: [para([run("Кафедра управления и информатики в технических системах", { size: 22 })], { alignment: AlignmentType.RIGHT, spacing: { line: 276 } })],
        }),
      ],
    })],
  }));

  children.push(...empty(2));
  children.push(para(
    [run(`Лабораторная работа №${LAB_NUM}`, { bold: true }), new TextRun({ break: 1, text: " ПО ДИСЦИПЛИНЕ", bold: true, size: BODY_SIZE, font: FONT })],
    { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } },
  ));
  children.push(para([run("«ИСИТ»")], { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } }));

  children.push(...empty(1));
  const studentGrid = [20, 8, 6, 6, 27, 14, 19];
  children.push(tumanTable([
    labelValueRow([
      { t: "СТУДЕНТА", span: 1 },
      { t: "3", span: 1, underline: true, align: AlignmentType.CENTER },
      { t: "КУРСА", span: 2 },
      { t: "БАКАЛАВРИАТА", span: 1, underline: true, align: AlignmentType.CENTER },
      { t: "ГРУППЫ", span: 1 },
      { t: "ИДБ-23-13", span: 1, underline: true, align: AlignmentType.CENTER },
    ], studentGrid),
    labelValueRow([
      { t: "", span: 3 },
      { t: "", span: 2, hint: "(уровень профессионального образования)" },
      { t: "", span: 2 },
    ], studentGrid),
  ]));

  children.push(...empty(1));
  children.push(tumanTable([
    labelValueRow([{ t: "Бектеев Георгий", underline: true, align: AlignmentType.CENTER }]),
    labelValueRow([{ t: "", hint: "(ФИО)" }]),
  ]));

  children.push(...empty(1));
  children.push(tumanTable([
    labelValueRow([{ t: "«C4 Model. Диаграмма кода (Level 4)»", underline: true, align: AlignmentType.CENTER }]),
    labelValueRow([{ t: "", hint: "(ТЕМА РАБОТЫ)" }]),
  ]));

  children.push(...empty(2));
  const dirGrid = [25, 75];
  children.push(tumanTable([
    labelValueRow([{ t: "Направление:" }, { t: "09.03.03 Прикладная информатика" }], dirGrid),
    labelValueRow([{ t: "Профиль подготовки:" }, { t: "«Управление данными»" }], dirGrid),
  ]));

  children.push(...empty(1));

  const signGrid = [22, 43, 15, 20];
  children.push(tumanTable([
    labelValueRow([{ t: "Отчет сдан «___» __________ 2026 г.", span: 4 }], signGrid),
    labelValueRow([{ t: "", span: 4 }], signGrid),
    labelValueRow([{ t: "Оценка ______________________________", span: 4 }], signGrid),
    labelValueRow([{ t: "", span: 4 }], signGrid),
    labelValueRow([
      { t: "Преподаватель" },
      { t: "Бычков Сергей Юрьевич", underline: true, align: AlignmentType.CENTER },
      { t: "" },
      { t: "", underline: true, align: AlignmentType.CENTER },
    ], signGrid),
    labelValueRow([
      { t: "" },
      { t: "", hint: "(Ф.И.О., должность, степень, звание.)" },
      { t: "" },
      { t: "", hint: "(подпись)" },
    ], signGrid),
  ]));

  children.push(...empty(2));
  children.push(para([run("МОСКВА 2026")], { alignment: AlignmentType.CENTER }));
  return children;
}

function buildContents() {
  const items = [
    "1. Цель работы",
    "2. Постановка задачи",
    "3. Исходные данные. Компонент «Сервис прогнозирования»",
    "4. Диаграмма классов (Level 4)",
    "5. Реализация прототипа на Python",
    "6. Результаты выполнения прототипа",
    "7. Глоссарий идентификаторов (Glossary IDs)",
    "8. Соответствие требованиям методического пособия",
    "9. Выводы",
  ];
  return [
    para([run("СОДЕРЖАНИЕ", { bold: true, size: H1_SIZE })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
    ...items.map(t => para([run(t, { size: H1_SIZE })], { spacing: { after: 120, line: LINE_BODY } })),
  ];
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 360, after: 200, line: LINE_BODY },
    children: [new TextRun({ text, bold: true, size: H1_SIZE, color: BLACK, font: FONT })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 280, after: 160, line: LINE_BODY },
    children: [new TextRun({ text, bold: true, size: H2_SIZE, color: BLACK, font: FONT })],
  });
}
function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: INDENT },
    spacing: { line: LINE_BODY, after: 80 },
    children: [new TextRun({ text, size: BODY_SIZE, color: BLACK, font: FONT })],
  });
}
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 240, line: 300 },
    children: [new TextRun({ text, size: BODY_SIZE, color: BLACK, font: FONT })],
  });
}
function tableTitle(text) {
  return new Paragraph({
    keepNext: true,
    alignment: AlignmentType.LEFT,
    spacing: { before: 160, after: 80, line: 300 },
    children: [new TextRun({ text, size: BODY_SIZE, bold: true, color: BLACK, font: FONT })],
  });
}
// Листинг кода: блок строк моноширинным шрифтом на светлом фоне
function codeBlock(lines) {
  return lines.map((line, i) => new Paragraph({
    keepNext: i < lines.length - 1 ? undefined : false,
    shading: { type: ShadingType.CLEAR, fill: "F6F8FA" },
    spacing: { line: 240, before: i === 0 ? 120 : 0, after: i === lines.length - 1 ? 240 : 0 },
    indent: { left: 227 },
    children: [new TextRun({ text: line.length ? line : " ", size: CODE_SIZE, color: BLACK, font: CODE_FONT })],
  }));
}
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function figure(file, widthPx = 615) {
  const { w, h } = pngSize(file);
  const height = Math.round(widthPx * h / w);
  return new Paragraph({
    keepNext: true,
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 40 },
    children: [new ImageRun({ data: fs.readFileSync(file), transformation: { width: widthPx, height }, type: "png" })],
  });
}
function dataTable(header, rows, widths) {
  const mkCell = (text, isHeader, widthPct) => new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: isHeader ? { type: ShadingType.CLEAR, fill: "F1F5F9" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { line: 276 },
      children: [new TextRun({ text, size: TABLE_SIZE, bold: isHeader, color: BLACK, font: FONT })],
    })],
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: BLACK },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: BLACK },
      left: { style: BorderStyle.SINGLE, size: 2, color: "808080" },
      right: { style: BorderStyle.SINGLE, size: 2, color: "808080" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "A0A0A0" },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "A0A0A0" },
    },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((t, i) => mkCell(t, true, widths[i])) }),
      ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => mkCell(t, false, widths[i])) })),
    ],
  });
}

// ── Данные отчёта ──
const imgClass = path.join(LAB, "diagrams/png/class_diagram.png");

const listingStrategy = codeBlock([
  'class ForecastStrategy(ABC):',
  '    """Интерфейс стратегии прогнозирования динамики населения."""',
  '',
  '    name: str = "strategy"',
  '',
  '    @abstractmethod',
  '    def forecast(self, series: TimeSeries,',
  '                 horizon: int = 10) -> ForecastResult:',
  '        """Строит прогноз на horizon лет по историческому ряду."""',
  '        raise NotImplementedError',
]);

const listingContext = codeBlock([
  'class ForecastContext:',
  '    """Контекст: делегирует расчёт прогноза активной стратегии."""',
  '',
  '    def __init__(self, strategy: ForecastStrategy) -> None:',
  '        self._strategy = strategy',
  '',
  '    @property',
  '    def strategy_name(self) -> str:',
  '        """Имя активного алгоритма (для логирования и UI)."""',
  '        return self._strategy.name',
  '',
  '    def set_strategy(self, strategy: ForecastStrategy) -> None:',
  '        """Динамическая подмена алгоритма во время выполнения."""',
  '        self._strategy = strategy',
  '',
  '    def run_forecast(self, series: TimeSeries,',
  '                     horizon: int = 10) -> ForecastResult:',
  '        """Единая точка входа: расчёт прогноза активной стратегией."""',
  '        return self._strategy.forecast(series, horizon)',
]);

const listingFit = codeBlock([
  'def _fit(self, y: list[float],',
  '         params: Params) -> tuple[list[float], float, float,',
  '                                  float | None]:',
  '    """Проход рекуррентных уравнений Хольта по ряду."""',
  '    alpha, beta, phi = (float(params["alpha"]),',
  '                        float(params["beta"]), float(params["phi"]))',
  '    multiplicative = params["trend"] == "mul"',
  '    level = y[0]',
  '    trend = ((y[1] - y[0]) / y[0]',
  '             if (multiplicative and y[0] != 0) else (y[1] - y[0]))',
  '    fitted: list[float] = [level]',
  '    sse = 0.0',
  '    for t in range(1, len(y)):',
  '        forecast_point = (level * (1.0 + phi * trend)',
  '                          if multiplicative else level + phi * trend)',
  '        error = y[t] - forecast_point',
  '        sse += error ** 2',
  '        if multiplicative and level != 0:',
  '            new_level = (alpha * y[t]',
  '                         + (1 - alpha) * level * (1.0 + phi * trend))',
  '            new_trend = (beta * (new_level / level - 1.0)',
  '                         + (1 - beta) * phi * trend)',
  '        else:',
  '            new_level = alpha * y[t] + (1 - alpha) * (level + phi * trend)',
  '            new_trend = beta * (new_level - level) + (1 - beta) * phi * trend',
  '        level, trend = new_level, new_trend',
  '        fitted.append(forecast_point)',
  '    return fitted, level, trend, sse',
]);

const listingDemo = codeBlock([
  '==============================================================================',
  'DemoScope: прогноз численности населения региона на 5 лет',
  'Исторический ряд: 2010-2025 (1553.9 -> 1437.0 тыс. чел.)',
  '==============================================================================',
  '',
  'Алгоритм: Хольт-Уинтерс (двойное экспоненциальное сглаживание) [holt_winters]',
  "  Параметры: {'trend': 'mul', 'damped': False, 'alpha': 0.8, 'beta': 0.5,",
  "  'phi': 1.0, 'aic': 42.3}",
  '  Метрики аппроксимации: MAPE=0.18% RMSE=3.21 MAE=2.75',
  '  Прогноз (тыс. чел.):',
  '    2026:   1428.5   (95% ДИ:   1422.2 ..   1434.8)',
  '    2027:   1420.3   (95% ДИ:   1411.4 ..   1429.2)',
  '    2028:   1412.0   (95% ДИ:   1401.1 ..   1422.9)',
  '    2029:   1403.8   (95% ДИ:   1391.2 ..   1416.4)',
  '    2030:   1395.5   (95% ДИ:   1381.4 ..   1409.6)',
  '',
  'Алгоритм: ARIMA(2,1,0) [arima]',
  "  Параметры: {'p': 2, 'd': 1, 'q': 0, 'aic': 2.3}",
  '  Метрики аппроксимации: MAPE=0.05% RMSE=1.08 MAE=0.72',
  '  Прогноз (тыс. чел.):',
  '    2026:   1428.6   (95% ДИ:   1426.7 ..   1430.4)',
  '    2027:   1419.8   (95% ДИ:   1417.2 ..   1422.4)',
  '    2028:   1410.8   (95% ДИ:   1407.5 ..   1414.0)',
  '    2029:   1401.6   (95% ДИ:   1397.9 ..   1405.3)',
  '    2030:   1392.4   (95% ДИ:   1388.2 ..   1396.5)',
  '',
  'Алгоритм: Когортно-компонентный метод [cohort]',
  "  Параметры: {'шаг модели': '1 год', 'СКР': 1.5,",
  "  'миграция, тыс./год': -2.78}",
  '  Метрики аппроксимации: MAPE=0.53% RMSE=9.37 MAE=7.93',
  '  Прогноз (тыс. чел.):',
  '    2026:   1423.4',
  '    2027:   1414.1',
  '    2028:   1404.5',
  '    2029:   1394.7',
  '    2030:   1384.7',
  '',
  '==============================================================================',
  'Проверка SOLID: контекст принял 3 стратегий(и) из реестра',
  'ALGORITHMS без изменения клиентского кода — принцип Open/Closed выполнен.',
  '==============================================================================',
]);

const bodyChildren = [
  h1("1. Цель работы"),
  body("Получить практические навыки детализации архитектуры на уровне кода: научиться отображать объектно-ориентированную структуру архитектурного компонента в виде диаграммы классов UML с интерфейсами, атрибутами, методами и типами связей, а также реализовывать спроектированную структуру в виде чистого объектно-ориентированного кода на Python с обязательным применением шаблона проектирования «Стратегия» (Strategy), строгой аннотацией типов (Type Hinting) и соблюдением принципов SOLID."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №3 необходимо:"),
  body("1) разработать диаграмму классов UML (UML Class Diagram) для выбранного компонента: отобразить структуру классов, интерфейсов и типов объектно-ориентированных связей (реализация, композиция, зависимость);"),
  body("2) реализовать прототип компонента на Python: чистый объектно-ориентированный код с обязательным использованием шаблона «Стратегия» для динамической подмены алгоритмов бизнес-логики, снабжённый строгой аннотацией типов и соответствующий принципам SOLID;"),
  body("3) выполнить прототип и зафиксировать результаты его работы;"),
  body("4) оформить отчёт по лабораторной работе."),

  h1("3. Исходные данные. Компонент «Сервис прогнозирования»"),
  body("Объектом проектирования является компонент «Сервис прогнозирования» (forecast_comp), выделенный в составе контейнера Backend API на диаграмме компонентов (лабораторная работа №2). Компонент выполняет оркестрацию прогнозирования динамики населения тремя алгоритмами: двойное экспоненциальное сглаживание (Хольт-Уинтерс), ARIMA и когортно-компонентный метод; для статистических моделей формируются доверительные интервалы и метрики аппроксимации (MAPE, RMSE, MAE), для когортно-компонентного метода — метрики аппроксимации по бэктесту."),
  body("Границы прототипа. Прототип сфокусирован на паттерне «Стратегия» и воспроизводит ядро компонента — выбор алгоритма и расчёт прогноза. Функциональность боевого контракта run_forecast из лабораторной работы №2, не относящаяся к подмене алгоритмов (сценарные поправки к ряду, работа с кэшем прогнозов через db_repo, приём когортных данных из cohort_data), остаётся в сервисном слое компонента и в прототип не включена: замена параметра model_type на динамическую подмену стратегии — предмет именно этой лабораторной работы."),
  body("Ключевое архитектурное требование к компоненту — динамическая подмена алгоритма прогнозирования без изменения клиентского кода: маршрутизатор REST API выбирает алгоритм параметром запроса, и добавление нового метода прогнозирования не должно затрагивать ни контекст выполнения, ни существующие реализации. Это требование напрямую диктует применение шаблона «Стратегия» и соответствует принципу открытости/закрытости (Open/Closed Principle) из SOLID."),

  h1("4. Диаграмма классов (Level 4)"),
  body("Диаграмма классов отображает объектно-ориентированную структуру компонента. Абстрактный интерфейс ForecastStrategy фиксирует единый контракт метода forecast(); три конкретные стратегии реализуют контракт (пунктирные стрелки с незакрашенным треугольником — отношение реализации); контекст ForecastContext владеет ссылкой на стратегию (ромб — агрегация «1 к 1») и делегирует ей расчёт. Контракты данных (TimeSeries, ForecastResult, ForecastMetrics) оформлены неизменяемыми data-классами."),
  figure(imgClass),
  caption("Рисунок 1 — Диаграмма классов компонента «Сервис прогнозирования»"),

  h2("4.1. Спецификация классов"),
  tableTitle("Таблица 1 — Классы прототипа"),
  dataTable(
    ["Класс", "Категория", "Ответственность"],
    [
      ["ForecastStrategy", "Абстрактный интерфейс", "Единый контракт прогнозирования: метод forecast(series, horizon) для всех алгоритмов"],
      ["HoltWintersStrategy", "Конкретная стратегия", "Двойное экспоненциальное сглаживание; сеточный подбор alpha, beta, phi по критерию AIC"],
      ["ArimaStrategy", "Конкретная стратегия", "ARIMA(p, d, 0): порядок d — по выраженности тренда, порядок p — сеточный поиск по AIC"],
      ["CohortComponent\u00ADStrategy", "Конкретная стратегия", "Когортно-компонентный метод: продвижение пирамиды населения с учётом рождаемости, смертности и миграции"],
      ["ForecastContext", "Контекст", "Хранение активной стратегии, динамическая подмена (set_strategy), единая точка входа run_forecast"],
      ["TimeSeries", "Контракт данных", "Неизменяемый исторический ряд (годы и значения)"],
      ["ForecastResult", "Контракт данных", "Единая структура результата: прогноз, доверительные интервалы, метрики, параметры модели"],
      ["ForecastMetrics", "Контракт данных", "Метрики качества аппроксимации (MAPE, RMSE, MAE)"],
    ],
    [24, 22, 54],
  ),

  h1("5. Реализация прототипа на Python"),

  h2("5.1. Структура модуля"),
  body("Прототип размещён в пакете forecasting; он воспроизводит набор алгоритмов и контракты боевого модуля проекта. Реализация выполнена на чистой стандартной библиотеке Python (без внешних зависимостей), что позволяет запускать прототип в любом окружении."),
  ...codeBlock([
    'лаб3/src/',
    '├── demo.py                     # демонстрация: подмена стратегий в runtime',
    '└── forecasting/',
    '    ├── __init__.py             # публичный API пакета',
    '    ├── models.py               # контракты данных (dataclasses)',
    '    ├── metrics.py              # метрики качества (mape, rmse, mae)',
    '    ├── base.py                 # абстрактный контракт стратегии (ABC)',
    '    ├── holt_winters.py         # стратегия Хольта-Уинтерса',
    '    ├── arima.py                # стратегия ARIMA(p,d,0)',
    '    ├── cohort_component.py     # когортно-компонентная стратегия',
    '    └── context.py              # контекст + реестр алгоритмов',
  ]),

  h2("5.2. Ключевые листинги"),
  body("Абстрактный контракт стратегии (base.py) — единая точка расширяемости компонента:"),
  ...listingStrategy,
  body("Контекст с динамической подменой стратегии (context.py) — свойство strategy_name даёт клиенту имя активного алгоритма:"),
  ...listingContext,
  body("Ядро алгоритма Хольта-Уинтерса (holt_winters.py) — рекуррентные уравнения сглаживания с демпфированием тренда:"),
  ...listingFit,

  h2("5.3. Соответствие принципам SOLID"),
  tableTitle("Таблица 2 — Соответствие принципам SOLID"),
  dataTable(
    ["Принцип", "Как соблюдается"],
    [
      ["S — Single Responsibility", "Каждый класс отвечает за одну задачу: стратегии — только расчёт прогноза, контекст — только выбор и делегирование, models — только контракты данных, metrics — только расчёт метрик"],
      ["O — Open/Closed", "Новый алгоритм добавляется классом-наследником ForecastStrategy и регистрацией в реестре; контекст и клиентский код не изменяются"],
      ["L — Liskov Substitution", "Все стратегии взаимозаменяемы: любая принимает TimeSeries и возвращает ForecastResult, что подтверждено демонстрацией подмены в runtime"],
      ["I — Interface Segregation", "Интерфейс стратегии содержит единственный метод forecast — клиенты не зависят от методов, которые не используют"],
      ["D — Dependency Inversion", "Контекст зависит от абстракции ForecastStrategy, а не от конкретных алгоритмов; конкретные стратегии внедряются в конструктор"],
    ],
    [26, 74],
  ),
  body("Требование Type Hinting выполнено полностью: сигнатуры всех публичных и приватных методов аннотированы (включая list[float], tuple[int, ...], dict[str, str | int | float], ForecastMetrics | None), данные оформлены frozen-data-классами."),

  h1("6. Результаты выполнения прототипа"),
  body("Прототип выполнен командой python demo.py. Демонстрационный сценарий прогоняет один и тот же ряд (численность населения депопулирующего региона, 2010–2025, тыс. чел.) через единый контекст; стратегии берутся из реестра ALGORITHMS, поэтому подключение нового алгоритма не потребует изменения клиентского кода:"),
  ...listingDemo,
  body("Интерпретация результатов:"),
  body("1. Хольт-Уинтерс и ARIMA дают согласованную траекторию (около минус 8–9 тыс. чел. в год) с малой ошибкой аппроксимации; модель Хольта-Уинтерса выбрала мультипликативный тренд (AIC=42.3, MAPE=0.18%), а ARIMA автоматически выбрала порядок интегрирования d=1 (выраженный линейный тренд) и порядок p=2 по критерию AIC (MAPE=0.05%)."),
  body("2. Когортно-компонентный метод показывает высокую точность бэктеста (MAPE=0.53%) и прогнозирует несколько более крутую убыль (минус 9,3–10,0 тыс. чел. в год): модель «старит» пирамиду населения, и доля пожилых когорт с высокой смертностью растёт — это демографически корректный эффект инерции, невидимый для статистических моделей по итоговому ряду. Доверительные интервалы для когортной модели не строятся: при фиксированных коэффициентах результат детерминирован, качество оценивается бэктестом."),
  body("3. Все три результата объединены единым контрактом ForecastResult: маршрутизатор REST API получает одинаковую структуру независимо от выбранного алгоритма, что подтверждает работоспособность паттерна «Стратегия»."),

  h1("7. Глоссарий идентификаторов (Glossary IDs)"),
  body("К идентификаторам уровней L1–L3 добавлены идентификаторы классов прототипа:"),
  tableTitle("Таблица 3 — Глоссарий идентификаторов классов"),
  dataTable(
    ["ID", "Элемент", "Тип"],
    [
      ["ForecastStrategy", "Интерфейс стратегии прогнозирования", "Абстракция (ABC)"],
      ["HoltWintersStrategy", "Стратегия Хольта-Уинтерса", "Конкретная стратегия"],
      ["ArimaStrategy", "Стратегия ARIMA(p,d,0)", "Конкретная стратегия"],
      ["CohortComponentStrategy", "Когортно-компонентная стратегия", "Конкретная стратегия"],
      ["ForecastContext", "Контекст прогнозирования", "Контекст (паттерн)"],
      ["TimeSeries, ForecastResult, ForecastMetrics", "Контракты данных", "Frozen dataclass"],
    ],
    [30, 42, 28],
  ),

  h1("8. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 4 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Диаграмма классов UML для выбранного компонента", "Выполнено: 8 классов с атрибутами, методами и типами связей (п. 4)"],
      ["Объектно-ориентированная структура: классы, интерфейсы, типы связей", "Выполнено: реализация (..|>), агрегация (o--), зависимости (..>)"],
      ["Обязательный паттерн «Стратегия»", "Выполнено: интерфейс + 3 конкретные стратегии + контекст с подменой в runtime"],
      ["Чистый ООП-код на Python", "Выполнено: пакет forecasting из 8 модулей, data-классы, ABC"],
      ["Строгая аннотация типов (Type Hinting)", "Выполнено: все сигнатуры аннотированы, включая обобщённые типы и объединения"],
      ["Соответствие принципам SOLID", "Выполнено: таблица соответствия в п. 5.3"],
      ["Прототип выполнен, результаты зафиксированы", "Выполнено: листинг вывода в п. 6"],
      ["Инструмент — PlantUML", "Выполнено: диаграмма классов описана текстом в PlantUML (Diagrams-as-Code)"],
    ],
    [55, 45],
  ),

  h1("9. Выводы"),
  body("В ходе лабораторной работы освоен четвёртый уровень абстракции модели C4 применительно к теме курсового проекта. Спроектирована диаграмма классов компонента «Сервис прогнозирования», зафиксировавшая объектно-ориентированную структуру: абстрактный контракт ForecastStrategy, три конкретные стратегии (Хольт-Уинтерс, ARIMA, когортно-компонентный метод), контекст с динамической подменой и неизменяемые контракты данных. Структура реализована в виде чистого объектно-ориентированного прототипа на Python с полной аннотацией типов и соблюдением всех пяти принципов SOLID; работоспособность подтверждена выполнением демонстрации, в которой один и тот же контекст последовательно принял три алгоритма без изменения клиентского кода. Полученный прототип воспроизводит логику боевого модуля проекта и образует основу для интеграции архитектурного описания в единый репозиторий (лабораторная работа №5)."),
];

const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Бектеев Георгий",
  title: "Отчёт по лабораторной работе №3 — C4 Model. Диаграмма кода",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: BODY_SIZE, color: BLACK },
        paragraph: { spacing: { line: LINE_BODY } },
      },
      heading1: {
        run: { font: FONT, size: H1_SIZE, bold: true, color: BLACK },
        paragraph: { spacing: { before: 360, after: 200, line: LINE_BODY } },
      },
      heading2: {
        run: { font: FONT, size: H2_SIZE, bold: true, color: BLACK },
        paragraph: { spacing: { before: 280, after: 160, line: LINE_BODY } },
      },
    },
  },
  sections: [
    { properties: { page: PAGE }, children: buildTitlePage() },
    { properties: { page: PAGE }, footers: { default: pageFooter() }, children: buildContents() },
    { properties: { page: PAGE }, footers: { default: pageFooter() }, children: bodyChildren },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("OK:", OUT, buf.length, "bytes");
});
