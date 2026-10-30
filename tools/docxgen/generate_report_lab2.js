// Генератор отчёта по ЛР2 (docx-js) — клон generate_report.js (титульник и оформление — по образцу пользователя)
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб2";
const OUT = path.join(LAB, "отчет_ЛР2.docx");
const LAB_NUM = "2";

// ── Оформление (по примеру) ──
const FONT = "Times New Roman";
const BODY_SIZE = 24;   // 12pt
const TABLE_SIZE = 24;  // 12pt
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
    labelValueRow([{ t: "«C4 Model. Диаграмма компонентов (Level 3)»", underline: true, align: AlignmentType.CENTER }]),
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
    "3. Исходные данные. Ключевой контейнер Backend API",
    "4. Диаграмма компонентов (Level 3)",
    "5. Спецификация внутренних интерфейсов",
    "6. Обоснование архитектурных решений",
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
const imgL3 = path.join(LAB, "diagrams/png/components_l3.png");

const bodyChildren = [
  h1("1. Цель работы"),
  body("Получить практические навыки декомпозиции контейнера на внутренние логические элементы с использованием третьего уровня абстракции методологии C4 Model: научиться выделять компоненты ключевого контейнера бизнес-логики, фиксировать их ответственность и внутренние интерфейсы, демонстрировать разделение ответственности в соответствии с принципами многослойной архитектуры. Закрепить подход Diagrams-as-Code — описание архитектуры в текстовом формате PlantUML с библиотекой C4-PlantUML."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №2 необходимо:"),
  body("1) разработать диаграмму компонентов (Component Diagram) для ключевого контейнера: выполнить декомпозицию контейнера бизнес-логики (бэкенда) на внутренние логические компоненты, выполняющиеся в едином адресном пространстве; диаграмма должна содержать не менее четырёх взаимосвязанных компонентов; продемонстрировать явное разделение ответственности между компонентами в соответствии с принципами многослойной архитектуры;"),
  body("2) подготовить текстовое обоснование принятых архитектурных решений и спецификацию внутренних интерфейсов компонентов;"),
  body("3) оформить отчёт по лабораторной работе."),

  h1("3. Исходные данные. Ключевой контейнер Backend API"),
  body("Объектом декомпозиции является контейнер api_backend (Backend API) системы DemoScope — платформы демографической аналитики, описанной в лабораторной работе №1. Контейнер реализован на стеке Python 3.10 / FastAPI с библиотеками Pandas, PyArrow и Statsmodels и исполняет ядро бизнес-логики системы:"),
  body("– REST API мониторинга демографических показателей (численность населения, рождаемость, смертность, миграция, заработная плата), тепловых карт и сравнения территорий;"),
  body("– прогнозирование динамики населения методами Хольта-Уинтерса, ARIMA и когортно-компонентным методом со сценарными поправками и доверительными интервалами;"),
  body("– выявление проблемных территорий («радар проблем») на основе предиктивных алертов;"),
  body("– генерация аналитических справок, в том числе с привлечением внешней большой языковой модели Google Gemini."),
  body("Backend API выбран в качестве ключевого контейнера обоснованно: именно в нём сосредоточена вся бизнес-логика системы, тогда как остальные контейнеры — клиентские приложения (web_app, mobile_app), хранилища (data_store, db_service) и конвейер подготовки данных (data_pipeline) — выполняют функции интерфейса, хранения и загрузки данных. В рамках распределения задач проектной команды №18 группы ИДБ-23-13 (Бектеев Г.А., Алибеков А., Лиджиев Э.С., Слисаренко Д., Сырлыбаев Б.С.) декомпозиция клиентского контейнера web_app параллельно выполняется фронтенд-разработчиком команды (Слисаренко Д.), а в настоящей работе детально рассматривается серверный контейнер api_backend. Внутри контейнера компоненты выполняются в едином адресном пространстве процесса и взаимодействуют прямыми программными вызовами (Python Call), без сетевых протоколов."),

  h1("4. Диаграмма компонентов (Level 3)"),
  body("Диаграмма компонентов выполняет «зум» внутрь контейнера api_backend: граница Container_Boundary объединяет внутренние логические компоненты, а внешние по отношению к бэкенду контейнеры (веб- и мобильное приложение, хранилища данных) и внешние системы вынесены за рамки границы. Внутри границы выделено шесть компонентов: маршрутизатор запросов, три сервиса бизнес-логики и два репозитория данных."),
  figure(imgL3),
  caption("Рисунок 1 — Диаграмма компонентов контейнера Backend API системы DemoScope"),

  h2("4.1. Спецификация компонентов"),
  tableTitle("Таблица 1 — Компоненты контейнера Backend API"),
  dataTable(
    ["Компонент", "ID", "Технологии", "Ответственность"],
    [
      ["API Router", "api_router", "FastAPI Router, Pydantic", "Маршрутизация REST-запросов (/api/population, /api/forecast, /api/report и др.), валидация входных схем, CORS-политика"],
      ["Сервис про\u00ADгнозирования", "forecast_comp", "Statsmodels", "Оркестрация прогнозов: выбор алгоритма (Хольт-Уинтерс, ARIMA, когортно-компонентный), сценарные поправки, доверительные интервалы"],
      ["Сервис аналитики", "analytics_comp", "Pandas", "Агрегация показателей, тепловые карты, сравнение территорий, расчёт топ-изменений"],
      ["Сервис отчётов и алертов", "report_comp", "Python Module", "Генерация аналитических справок (вызов LLM), предиктивные алерты «радара проблем»"],
      ["Репозиторий данных", "data_repo", "PyArrow / Parquet", "Единый слой доступа к озеру данных: чтение временных рядов и агрегатов, предрасчёт и кэширование"],
      ["Репозиторий служебной БД", "db_repo", "SQLAlchemy", "Доступ к служебной БД: учётные записи, сессии, история отчётов, кэш прогнозов"],
    ],
    [19, 20, 18, 43],
  ),

  h1("5. Спецификация внутренних интерфейсов"),
  body("Внутрипроцессное взаимодействие компонентов строится по принципу «слои через интерфейсы»: маршрутизатор обращается только к сервисам бизнес-логики, сервисы — только к репозиториям, репозитории — только к хранилищам. Ниже зафиксированы ключевые интерфейсы компонентов."),

  h2("5.1. Внешний интерфейс (REST API маршрутизатора)"),
  tableTitle("Таблица 2 — Операции REST API"),
  dataTable(
    ["Операция", "Метод и путь", "Назначение"],
    [
      ["Список территорий", "GET /api/regions, GET /api/municipalities", "Справочники субъектов и муниципалитетов"],
      ["Показатели мониторинга", "GET /api/population, GET /api/demographics, GET /api/salary", "Временные ряды показателей по территориям"],
      ["Аналитические выборки", "GET /api/top-changes, GET /api/heatmap, GET /api/compare/territories", "Топ-изменения, тепловая карта, сравнение территорий"],
      ["Про\u00ADгно\u00ADзи\u00ADро\u00ADва\u00ADние", "POST /api/forecast", "Прогноз динамики населения по выбранному алгоритму и сценарию"],
      ["Отчёты и алерты", "POST /api/report, GET /api/predictive-alerts, POST /api/llm-analysis", "Аналитические справки, «радар проблем», LLM-анализ"],
    ],
    [20, 36, 44],
  ),

  h2("5.2. Внутренние программные интерфейсы (Python Call)"),
  tableTitle("Таблица 3 — Внутренние программные интерфейсы компонентов"),
  dataTable(
    ["Владелец", "Интерфейс (сигнатура)", "Вызывающие компоненты"],
    [
      ["forecast_comp", "run_forecast(years, values, horizon, model_type, scenario, cohort_data) -> dict", "api_router"],
      ["analytics_comp", "get_heatmap_data(), get_compare_territories(region, years), get_top_changes(limit, years)", "api_router"],
      ["report_comp", "generate_report(region, indicators) -> Report, get_predictive_alerts() -> list[Alert]", "api_router"],
      ["data_repo", "get_population_data(region, years), get_population_timeseries(region), load_births(), load_deaths()", "forecast_comp, analytics_comp, report_comp"],
      ["db_repo", "get_forecast_cache(key), save_forecast_cache(key, payload), save_report_history(report)", "forecast_comp, report_comp"],
    ],
    [19, 51, 30],
  ),
  body("Формулировки сигнатур фиксируют целевые проектные контракты компонента (подход as-is/to-be, принятый в лабораторной работе №1): названия и состав операций сверены с действующей кодовой базой проекта, однако часть контрактов — в частности, доступ к кэшу прогнозов и истории отчётов через репозиторий db_repo, а также сценарные поправки в run_forecast — реализуется по мере развития проекта в спринтах."),

  h2("5.3. Спецификация связей"),
  tableTitle("Таблица 4 — Связи диаграммы компонентов"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Веб-приложение (SPA)", "api_router", "Запросы к REST API", "HTTPS/JSON"],
      ["Мобильное приложение", "api_router", "Запросы к REST API", "HTTPS/JSON"],
      ["api_router", "analytics_comp", "Вызов функций агрегации", "Python Call"],
      ["api_router", "forecast_comp", "Вызов точки входа прогнозов", "Python Call"],
      ["api_router", "report_comp", "Вызов генерации справок", "Python Call"],
      ["forecast_comp", "data_repo", "Чтение временных рядов", "Python Call"],
      ["analytics_comp", "data_repo", "Чтение показателей", "Python Call"],
      ["report_comp", "data_repo", "Чтение данных для справок", "Python Call"],
      ["forecast_comp", "db_repo", "Кэш прогнозов (чтение/запись)", "Python Call"],
      ["report_comp", "db_repo", "Сохранение истории отчётов", "Python Call"],
      ["data_repo", "Хранилище демографических данных", "Чтение Parquet-фрагментов", "Файловый I/O (Parquet)"],
      ["db_repo", "Служебная БД", "Выполнение SQL-запросов", "SQL/TCP 5432"],
      ["report_comp", "Google Gemini API", "Запрос генерации текста", "HTTPS/JSON (REST)"],
    ],
    [19, 21, 38, 22],
  ),
  body("Стрелки между внутренними компонентами — прямые программные вызовы в общем адресном пространстве. Изнутри контейнера наружу ведут только стрелки от репозиториев к хранилищам и от сервиса отчётов к внешнему LLM-сервису; входящие сетевые стрелки — запросы клиентских приложений к маршрутизатору."),

  h1("6. Обоснование архитектурных решений"),
  body("1. Многослойная архитектура. Компоненты сгруппированы в три слоя: слой представления (api_router), слой бизнес-логики (forecast_comp, analytics_comp, report_comp) и слой доступа к данным (data_repo, db_repo). Зависимости направлены строго вниз: маршрутизатор не обращается к хранилищам напрямую, а сервисы не знают о транспортном уровне. Такая структура изолирует замену любых горизонталей — например, смена веб-фреймворка не затрагивает бизнес-логику."),
  body("2. Выделение единого слоя доступа к данным. Всё чтение демографических массивов инкапсулировано в репозитории data_repo: сервисы работают с рядами и агрегатами, не зная деталей физического формата (Parquet) и схемы разбиения файлов озера данных. Это соответствует принципу инверсии зависимостей — бизнес-слой зависит от интерфейса репозитория, а не от технологии хранения."),
  body("3. Разделение аналитической и служебной нагрузки. Прогнозирование, аналитика и отчёты выделены в три независимых сервиса бизнес-логики: у каждого своя зона ответственности и собственные источники данных. Кэш прогнозов и история отчётов обслуживаются отдельным репозиторием db_repo с транзакционной СУБД, что отделяет служебную нагрузку от больших аналитических выборок."),
  body("4. Единая точка входа прогнозирования. Сервис прогнозирования предоставляет один интерфейс run_forecast с параметром выбора алгоритма: маршрутизатор не знает о трёх методах прогнозирования, их перебор и сценарные поправки инкапсулированы внутри компонента (детализация на уровне классов выполнена в лабораторной работе №3)."),
  body("5. Внешние интеграции за границей контейнера. Генерация текста справок выполняется внешним сервисом Google Gemini: компонент report_comp обращается к нему по HTTPS/JSON, что отражено на диаграмме выносом LLM-сервиса за границу Container_Boundary."),

  h1("7. Глоссарий идентификаторов (Glossary IDs)"),
  body("К идентификаторам уровней L1/L2 (см. лабораторную работу №1) добавлены идентификаторы компонентов бэкенда:"),
  tableTitle("Таблица 5 — Глоссарий идентификаторов компонентов"),
  dataTable(
    ["ID", "Элемент", "Тип"],
    [
      ["api_router", "API Router (маршрутизатор)", "Компонент"],
      ["forecast_comp", "Сервис прогнозирования", "Компонент"],
      ["analytics_comp", "Сервис аналитики", "Компонент"],
      ["report_comp", "Сервис отчётов и алертов", "Компонент"],
      ["data_repo", "Репозиторий данных (озеро данных)", "Компонент"],
      ["db_repo", "Репозиторий служебной БД", "Компонент"],
      ["web_app, mobile_app, data_store, db_service, gemini_api", "Внешние по отношению к бэкенду элементы", "Контейнеры / внешние системы (см. ЛР1)"],
    ],
    [26, 44, 30],
  ),

  h1("8. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 6 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Декомпозиция ключевого контейнера бизнес-логики на внутренние компоненты", "Выполнено: контейнер api_backend декомпозирован на 6 компонентов"],
      ["Не менее 4 взаимосвязанных компонентов в едином адресном пространстве", "Выполнено: 6 компонентов, взаимодействие — прямые программные вызовы"],
      ["Явное разделение ответственности (многослойная архитектура)", "Выполнено: 3 слоя — представление, бизнес-логика, доступ к данным (п. 6)"],
      ["Текстовое обоснование архитектурных решений", "Выполнено: п. 6, пять обоснований"],
      ["Спецификация внутренних интерфейсов", "Выполнено: п. 5.1 (REST), п. 5.2 (программные интерфейсы), п. 5.3 (связи)"],
      ["Протоколы на всех внешних связях компонента", "Выполнено: HTTPS/JSON, SQL/TCP 5432, файловый I/O, HTTPS/JSON (REST)"],
      ["Инструмент — PlantUML + C4-PlantUML", "Выполнено: макросы Component, Container_Boundary, Container, ContainerDb, System_Ext, Rel, LAYOUT_WITH_LEGEND"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: нарушений нет"],
    ],
    [55, 45],
  ),

  h1("9. Выводы"),
  body("В ходе лабораторной работы освоен третий уровень абстракции модели C4 применительно к теме курсового проекта. Ключевой контейнер системы DemoScope — Backend API — декомпозирован на шесть внутренних компонентов, сгруппированных в три слоя: маршрутизатор запросов, три сервиса бизнес-логики (прогнозирование, аналитика, отчёты и алерты) и два репозитория данных. Зафиксированы ответственность каждого компонента и его внутренние интерфейсы: внешний REST API, программные интерфейсы сервисов и репозиториев, протоколы всех внешних связей. Диаграмма построена в нотации C4-PlantUML в подходе Diagrams-as-Code. Выделение слоя доступа к данным и единой точки входа прогнозирования подготовило основу для следующего уровня детализации — проектирования классов прогнозного компонента с паттерном «Стратегия» (лабораторная работа №3)."),
];

const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Бектеев Георгий",
  title: "Отчёт по лабораторной работе №2 — C4 Model. Диаграмма компонентов",
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
