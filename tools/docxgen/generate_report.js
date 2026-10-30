// Генератор отчёта по ЛР1 (docx-js)
// Титульник и оформление — по образцу пользователя (пример.docx, шаблон СТАНКИН):
//   потоковый титул (шапка вуза -> Институт/Кафедра -> ЛАБОРАТОРНАЯ РАБОТА -> таблицы
//   «студента…»/ФИО/Направление -> Отчет сдан/Оценка/Преподаватель -> МОСКВА ГОД),
//   затем СОДЕРЖАНИЕ, затем основной текст. Базовый шрифт TNR 12pt, заголовки 14/13pt.
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб1";
const OUT = path.join(LAB, "отчет_ЛР1.docx");

// ── Оформление (по примеру) ──
const FONT = "Times New Roman";
const BODY_SIZE = 24;   // 12pt — базовый размер примера (Normal)
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

// ── Компоненты титульника ──
function labelValueRow(cells, widths) {
  // cells: [{t, span, underline, hint, align}] — строка таблицы титульника
  // widths: проценты сетки, накапливаются по span
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

  // Шапка вуза (5 строк, по центру, жирно — по шаблону СТАНКИН)
  children.push(para([]));
  [
    "МИНОБРНАУКИ РОССИИ",
    "федеральное государственное автономное образовательное учреждение",
    "высшего образования",
    "«Московский государственный технологический университет «СТАНКИН»",
    "(ФГАОУ ВО «МГТУ «СТАНКИН»)",
  ].forEach(t => children.push(para([run(t, { bold: true })], { alignment: AlignmentType.CENTER, spacing: { line: 276 } })));

  children.push(...empty(1));

  // Институт | Кафедра (11pt, по краям)
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

  // Название работы
  children.push(para(
    [run("Лабораторная работа №1", { bold: true }), new TextRun({ break: 1, text: " ПО ДИСЦИПЛИНЕ", bold: true, size: BODY_SIZE, font: FONT })],
    { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } },
  ));
  children.push(para([run("«ИСИТ»")], { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } }));

  children.push(...empty(1));

  // Таблица «студента … курса … группы …»
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

  // ФИО (ячейка с подчёркиванием + подпись 8pt)
  children.push(tumanTable([
    labelValueRow([{ t: "Бектеев Георгий", underline: true, align: AlignmentType.CENTER }]),
    labelValueRow([{ t: "", hint: "(ФИО)" }]),
  ]));

  children.push(...empty(1));

  // ТЕМА РАБОТЫ
  children.push(tumanTable([
    labelValueRow([{ t: "«C4 Model. Диаграмма системного контекста (Level 1) и диаграмма контейнеров (Level 2)»", underline: true, align: AlignmentType.CENTER }]),
    labelValueRow([{ t: "", hint: "(ТЕМА РАБОТЫ)" }]),
  ]));

  children.push(...empty(2));

  // Направление / Профиль
  const dirGrid = [25, 75];
  children.push(tumanTable([
    labelValueRow([{ t: "Направление:" }, { t: "09.03.03 Прикладная информатика" }], dirGrid),
    labelValueRow([{ t: "Профиль подготовки:" }, { t: "«Управление данными»" }], dirGrid),
  ]));

  children.push(...empty(1));

  // Отчет сдан / Оценка / Преподаватель
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

  children.push(...empty(1));
  children.push(para([run("МОСКВА 2026")], { alignment: AlignmentType.CENTER }));
  return children;
}

// ── СОДЕРЖАНИЕ (как в примере: заголовок + список разделов 14pt) ──
function buildContents() {
  const items = [
    "1. Цель работы",
    "2. Постановка задачи",
    "3. Исходные данные. Общая характеристика системы",
    "4. Диаграмма системного контекста (Level 1)",
    "5. Диаграмма контейнеров (Level 2)",
    "6. Глоссарий идентификаторов (Glossary IDs)",
    "7. Соответствие требованиям методического пособия",
    "8. Выводы",
  ];
  return [
    para([run("СОДЕРЖАНИЕ", { bold: true, size: H1_SIZE })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
    ...items.map(t => para([run(t, { size: H1_SIZE })], { spacing: { after: 120, line: LINE_BODY } })),
  ];
}

// ── Компоненты основного текста ──
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
const imgL1 = path.join(LAB, "diagrams/png/context_l1.png");
const imgL2 = path.join(LAB, "diagrams/png/containers_l2.png");

const bodyChildren = [
  h1("1. Цель работы"),
  body("Получить практические навыки проектирования архитектуры программной системы с использованием методологии C4 Model: научиться очерчивать жёсткие внешние границы проектируемой системы на диаграмме системного контекста (Level 1) и выполнять декомпозицию системы на независимо развёртываемые технологические контейнеры на диаграмме контейнеров (Level 2). Освоить подход Diagrams-as-Code — описание архитектуры в текстовом формате PlantUML с библиотекой C4-PlantUML."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №1 необходимо:"),
  body("1) разработать диаграмму системного контекста (System Context Diagram), в которой проектируемая система представлена единым неделимым «чёрным ящиком»; диаграмма должна содержать центральный блок, не менее трёх пользовательских ролей (персон) и не менее трёх внешних программных или аппаратных систем, находящихся вне зоны прямого контроля разработчиков; все связи должны быть строго направленными и снабжены текстовым описанием передаваемых данных;"),
  body("2) разработать диаграмму контейнеров (Container Diagram), выполнив декомпозицию центрального блока на обособленные, независимо развёртываемые исполняемые сущности; выделить не менее пяти технологических контейнеров (включая мобильное и веб-приложение, бэкенд-сервисы и хранилища данных); для каждого контейнера зафиксировать технологический стек; для всех межсервисных связей указать сетевые протоколы; сформировать таблицу описания контейнеров;"),
  body("3) оформить отчёт по лабораторной работе."),

  h1("3. Исходные данные. Общая характеристика системы"),
  body("Объектом архитектурного моделирования является информационная система DemoScope — платформа демографической аналитики, предназначенная для мониторинга, анализа и прогнозирования демографической ситуации в субъектах Российской Федерации. Система решает следующие задачи:"),
  body("– консолидация демографических и социально-экономических показателей Росстата и ЕМИСС (численность населения, рождаемость, смертность, миграция, заработная плата) в едином хранилище;"),
  body("– интерактивная визуализация данных (тепловые карты, 3D-глобус, динамические графики) и сравнение территорий;"),
  body("– прогнозирование динамики населения методами Хольта-Уинтерса, ARIMA и когортно-компонентным методом с расчётом доверительных интервалов;"),
  body("– выявление проблемных территорий («радар проблем») на основе предиктивных алертов;"),
  body("– автоматическая генерация аналитических справок, в том числе с использованием большой языковой модели (Google Gemini)."),
  body("Проект разрабатывается проектной командой №18 группы ИДБ-23-13 в следующем составе: Бектеев Г.А. (тимлид — координация архитектуры, интеграция контуров), Алибеков А. (back-end — серверное ядро, расчётные алгоритмы прогнозирования), Лиджиев Э.С. (tech-doc — спецификация API, документация), Слисаренко Д. (front-end — клиентское веб-приложение и визуализация), Сырлыбаев Б.С. (аналитик — сбор данных, конвейер ETL, аналитические отчёты)."),
  body("Курсовой проект находится на этапе архитектурного проектирования: состав подсистем задаётся целевым образом, а реализация ведётся поэтапно в рамках спринтов, поэтому отдельные компоненты (в частности, мобильное приложение) находятся в стадии разработки. Целевая архитектура проектируется на стеке Python/FastAPI (серверная часть) и React (клиентская часть); демографические массивы хранятся в озере данных формата Parquet, служебные данные — в СУБД PostgreSQL. Архитектурное описание выполняется в подходе Diagrams-as-Code в нотации C4-PlantUML."),

  h1("4. Диаграмма системного контекста (Level 1)"),
  body("Диаграмма системного контекста показывает проектируемую систему в её окружении с «высоты птичьего полёта»: система представлена единым неделимым «чёрным ящиком», внутреннее устройство на данном уровне абстракции не раскрывается. Диаграмма отвечает на вопрос, какую ценность система предоставляет пользователям и с какими внешними системами она взаимодействует."),
  figure(imgL1),
  caption("Рисунок 1 — Диаграмма системного контекста системы DemoScope"),

  h2("4.1. Спецификация элементов"),
  tableTitle("Таблица 1 — Элементы диаграммы системного контекста"),
  dataTable(
    ["Элемент", "Тип (C4)", "Описание"],
    [
      ["Демограф-аналитик", "Person", "Основной пользователь системы: мониторит демографические показатели, сравнивает территории, строит прогнозы"],
      ["Сотрудник органа государственного управления", "Person", "Использует «радар проблем» и аналитические отчёты для принятия решений в области социально-демографической политики"],
      ["Исследователь / студент", "Person", "Изучает демографическую динамику, выгружает справки и отчёты"],
      ["DemoScope", "Software System", "Проектируемая система: мониторинг, анализ и прогнозирование динамики населения РФ"],
      ["Росстат", "External System", "Источник официальной демографической статистики"],
      ["ЕМИСС", "External System", "Государственная платформа статистических данных"],
      ["Google Gemini API", "External System", "LLM-сервис генерации аналитических заключений (Gemini 2.5 Flash)"],
      ["Google Drive", "External System", "Хранилище исходных массивов данных (CSV), используемое конвейером загрузки"],
    ],
    [28, 18, 54],
  ),

  h2("4.2. Спецификация связей"),
  tableTitle("Таблица 2 — Связи диаграммы системного контекста"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "DemoScope", "Запросы мониторинга показателей, сравнения территорий, построения прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "DemoScope", "Запросы «радара проблем» и аналитических отчётов", "HTTPS"],
      ["Исследователь / студент", "DemoScope", "Запросы на изучение динамики, выгрузку справок и отчётов", "HTTPS"],
      ["DemoScope", "Росстат", "Запросы официальных демографических показателей", "HTTPS/REST"],
      ["DemoScope", "ЕМИСС", "Запросы статистических данных по регионам", "HTTPS/REST"],
      ["DemoScope", "Google Gemini API", "Запрос на генерацию текста аналитического заключения", "HTTPS/JSON"],
      ["DemoScope", "Google Drive", "Автозагрузка исходных массивов данных (CSV)", "HTTPS"],
    ],
    [22, 18, 42, 18],
  ),
  body("Все связи однонаправленные: направление стрелки отражает инициатора запроса и направление движения данных. Система является «чёрным ящиком»: внутренние контейнеры и компоненты на данном уровне не показываются."),

  h1("5. Диаграмма контейнеров (Level 2)"),
  body("Диаграмма контейнеров раскрывает внутреннюю структуру «чёрного ящика»: центральный блок декомпозирован на обособленные, независимо развёртываемые исполняемые сущности. Контейнер в терминологии C4 — не Docker-контейнер, а приложение или хранилище данных, которое должно выполняться для работы всей системы. Для каждого контейнера зафиксирован конкретный технологический стек, для всех межсервисных связей указаны протоколы взаимодействия."),
  figure(imgL2),
  caption("Рисунок 2 — Диаграмма контейнеров системы DemoScope"),

  h2("5.1. Таблица описания контейнеров"),
  tableTitle("Таблица 3 — Технологические контейнеры системы"),
  dataTable(
    ["ID", "Контейнер", "Тип", "Тех\u00ADно\u00ADло\u00ADги\u00ADче\u00ADский стек", "Назначение"],
    [
      ["web_app", "Веб-приложение (SPA)", "Container", "React 19 / Vite, Three.js, Recharts, Leaflet", "Основной интерфейс: 3D-глобус, тепловые карты, дашборды, графики, сравнение территорий"],
      ["mobile_app", "Мобильное приложение", "Container", "React Native", "Доступ к мониторингу и прогнозам со смартфона через единый REST API backend-контейнера"],
      ["api_backend", "Backend API", "Container", "Python 3.10 / FastAPI, Pandas, PyArrow, Statsmodels", "Исполняемое ядро бизнес-логики: REST API мониторинга, сравнения территорий, про\u00ADгно\u00ADзи\u00ADро\u00ADва\u00ADния, генерации отчётов, LLM-анализа"],
      ["data_pipeline", "Конвейер подготовки данных", "Container", "Python / Pandas, gdown", "Служебные утилиты: загрузка исходных массивов, очистка, конвертация CSV в Parquet, геомаппинг регионов"],
      ["data_store", "Хранилище де\u00ADмо\u00ADгра\u00ADфи\u00ADче\u00ADских данных", "ContainerDb", "Parquet / PyArrow (озеро данных)", "Хранение показателей Росстата и ЕМИСС: население, рождаемость, смертность, миграция (~6 ГБ)"],
      ["db_service", "Служебная БД", "ContainerDb", "PostgreSQL 16", "Учётные записи и сессии пользователей, история сформированных отчётов, кэш результатов прогнозов"],
    ],
    [22, 19, 17, 21, 21],
  ),

  h2("5.2. Спецификация связей"),
  tableTitle("Таблица 4 — Связи диаграммы контейнеров"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "web_app", "Работа с дашбордами и прогнозами", "HTTPS"],
      ["Сотрудник органа гос. управления", "web_app", "Запросы отчётов и «радара проблем»", "HTTPS"],
      ["Сотрудник органа гос. управления", "mobile_app", "Оперативный просмотр показателей и алертов", "HTTPS"],
      ["Исследователь / студент", "web_app", "Изучение демографической динамики, выгрузка справок", "HTTPS"],
      ["web_app", "api_backend", "Запросы к REST API (/api/population, /api/forecast, /api/report и др.)", "HTTPS/JSON"],
      ["mobile_app", "api_backend", "Запросы к REST API", "HTTPS/JSON"],
      ["api_backend", "data_store", "Чтение показателей и агрегатов", "Файловый I/O (Parquet)"],
      ["api_backend", "db_service", "Чтение и запись служебных данных", "SQL/TCP 5432"],
      ["data_pipeline", "data_store", "Запись подготовленных данных", "Файловый I/O (Parquet)"],
      ["data_pipeline", "Росстат", "Загрузка официальных демографических показателей", "HTTPS/REST"],
      ["data_pipeline", "ЕМИСС", "Загрузка региональных статистических витрин", "HTTPS/REST"],
      ["data_pipeline", "Google Drive", "Скачивание исходных массивов (CSV)", "HTTPS (gdown)"],
      ["api_backend", "Google Gemini API", "Генерация аналитического заключения", "HTTPS/JSON (REST)"],
    ],
    [22, 18, 42, 18],
  ),

  h2("5.3. Обоснование архитектурных решений"),
  body("1. Разделение интерфейсного и сервисного слоёв. Веб-приложение вынесено в отдельный контейнер: оно независимо развёртывается (статическая сборка Vite) и обращается к бэкенду только через REST API, что обеспечивает слабую связанность клиентского и серверного слоёв."),
  body("2. Отдельный контейнер конвейера данных. Подготовка данных (загрузка, очистка, конвертация) выполняется офлайн-утилитами, не входящими в исполняемый процесс API: конвейер может запускаться по расписанию независимо от работы сервиса, что соответствует принципу единственной ответственности на уровне развёртывания."),
  body("3. Двухуровневая модель данных. Аналитическая нагрузка системы — редкие чтения больших неизменяемых массивов — обслуживается озером данных на колоночном формате Parquet, обрабатываемым PyArrow «в памяти»: это обеспечивает высокую производительность аналитических выборок без администрирования тяжёлой СУБД. Служебные же данные (учётные записи, сессии, история отчётов, кэш прогнозов) хранятся в PostgreSQL — транзакционность гарантирована, а межсервисная связь с БД выполняется по стандартному сетевому протоколу SQL/TCP 5432."),
  body("4. Мобильное приложение в контуре системы. Курсовой проект находится на этапе архитектурного проектирования: состав контейнеров задаётся целевым образом и наполняется реализацией поэтапно (спринты проекта). Мобильное приложение включено в архитектуру наравне с веб-приложением и разделяет с ним единый REST API: стек React Native переиспользует компетенции команды (React 19) и не влечёт изменений в остальных контейнерах."),
  body("5. Внешние LLM-интеграции за границей системы. Генерация аналитических заключений выполняется внешним сервисом Google Gemini: система не содержит собственной модели, что отражено выносом API за границу System_Boundary."),

  h1("6. Глоссарий идентификаторов (Glossary IDs)"),
  body("Единые текстовые идентификаторы элементов, используемые во всех диаграммах проекта:"),
  tableTitle("Таблица 5 — Глоссарий идентификаторов"),
  dataTable(
    ["ID", "Элемент", "Тип"],
    [
      ["analyst", "Демограф-аналитик", "Пользователь"],
      ["official", "Сотрудник органа гос. управления", "Пользователь"],
      ["researcher", "Исследователь / студент", "Пользователь"],
      ["demoscope", "DemoScope (граница системы)", "Software System"],
      ["web_app", "Веб-приложение (SPA)", "Контейнер"],
      ["mobile_app", "Мобильное приложение", "Контейнер"],
      ["api_backend", "Backend API", "Контейнер"],
      ["data_pipeline", "Конвейер подготовки данных", "Контейнер"],
      ["data_store", "Хранилище демографических данных", "Контейнер (БД)"],
      ["db_service", "Служебная БД (PostgreSQL)", "Контейнер (БД)"],
      ["rosstat", "Росстат", "Внешняя система"],
      ["emiss", "ЕМИСС", "Внешняя система"],
      ["gemini_api", "Google Gemini API", "Внешняя система"],
      ["gdrive", "Google Drive", "Внешняя система"],
    ],
    [22, 48, 30],
  ),

  h1("7. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 6 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Система — единый неделимый «чёрный ящик» на L1", "Выполнено: внутренняя структура на L1 не раскрывается"],
      ["Центральный блок и не менее 3 пользовательских ролей", "Выполнено: 3 персоны"],
      ["Не менее 3 внешних систем вне зоны контроля разработчиков", "Выполнено: 4 внешние системы"],
      ["Все связи строго направленные, с описанием данных", "Выполнено: 7 направленных связей с описаниями и протоколами"],
      ["Декомпозиция центрального блока на контейнеры (L2)", "Выполнено: 6 контейнеров"],
      ["Не менее 5 контейнеров, включая мобильное и веб-приложение, бэкенд, хранилища", "Выполнено: все обязательные категории представлены (веб- и мобильное приложение, Backend API, конвейер данных, озеро данных Parquet и служебная БД PostgreSQL)"],
      ["Технологический стек для каждого контейнера", "Выполнено: зафиксирован в таблице 3"],
      ["Сетевые протоколы для всех межсервисных связей", "Выполнено: указаны для всех 13 связей (HTTPS, HTTPS/REST, HTTPS/JSON, SQL/TCP 5432, файловый I/O озера данных)"],
      ["Таблица описания контейнеров", "Выполнено: п. 5.1"],
      ["Инструмент — PlantUML + C4-PlantUML", "Выполнено: макросы Person, System, System_Ext, Container, ContainerDb, Rel, LAYOUT_WITH_LEGEND, LAYOUT_LEFT_RIGHT"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: нарушений нет"],
    ],
    [55, 45],
  ),

  h1("8. Выводы"),
  body("В ходе лабораторной работы освоены первые два уровня абстракции модели C4 применительно к теме курсового проекта. Построена диаграмма системного контекста, очерчивающая жёсткие внешние границы системы DemoScope: три пользовательские роли, четыре внешние системы и семь направленных информационных потоков с указанием протоколов. Выполнена декомпозиция системы на шесть технологических контейнеров с фиксацией стека и протоколов тринадцати направленных межсервисных и внешних связей. Архитектурное описание оформлено в подходе Diagrams-as-Code: диаграммы описаны текстом в нотации C4-PlantUML, версионируются и автоматически рендерятся в графические изображения. Полученные диаграммы образуют базу для последующих уровней детализации: декомпозиции бэкенда на компоненты (Level 3) и проектирования классов (Level 4)."),
];

const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Бектеев Георгий",
  title: "Отчёт по лабораторной работе №1 — C4 Model",
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
