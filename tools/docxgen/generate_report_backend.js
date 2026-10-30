// Генератор отчёта по ЛР1 — вариант для участника с ролью бекендер (лаб1_бекендер).
// Титульник и оформление — те же, что в базовом отчёте (образец СТАНКИН);
// основной текст переписан другими формулировками, диаграммы — из лаб1_бекендер/diagrams.
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб1_бекендер";
const OUT = path.join(LAB, "отчет_ЛР1.docx");

// ── Оформление (как в базовом отчёте) ──
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

// ── Компоненты титульника ──
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

  // Шапка вуза
  children.push(para([]));
  [
    "МИНОБРНАУКИ РОССИИ",
    "федеральное государственное бюджетное образовательное учреждение",
    "высшего образования",
    "«Московский государственный технологический университет «СТАНКИН»",
    "(ФГБОУ ВО «МГТУ «СТАНКИН»)",
  ].forEach(t => children.push(para([run(t, { bold: true })], { alignment: AlignmentType.CENTER, spacing: { line: 276 } })));

  children.push(empty(1));

  // Институт | Кафедра
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

  children.push(empty(4));

  // Название работы
  children.push(para(
    [run("ЛАБОРАТОРНАЯ РАБОТА №1", { bold: true }), new TextRun({ break: 1, text: "ПО ДИСЦИПЛИНЕ", bold: true, size: BODY_SIZE, font: FONT })],
    { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } },
  ));
  children.push(para([run("«ИНФОРМАЦИОННЫЕ СИСТЕМЫ И ТЕХНОЛОГИИ»")], { alignment: AlignmentType.CENTER, spacing: { line: LINE_BODY } }));

  children.push(empty(1));

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

  children.push(empty(1));

  // ФИО
  children.push(tumanTable([
    labelValueRow([{ t: "", underline: true }]),
    labelValueRow([{ t: "", hint: "(ФИО)" }]),
  ]));

  children.push(empty(5));

  // Направление / Профиль
  const dirGrid = [25, 75];
  children.push(tumanTable([
    labelValueRow([{ t: "Направление:" }, { t: "09.03.03 Прикладная информатика" }], dirGrid),
    labelValueRow([{ t: "Профиль подготовки:" }, { t: "Управление данными" }], dirGrid),
  ]));

  children.push(empty(1));
  children.push(para([run("Вариант работы: №4")], { spacing: { line: 276 } }));
  children.push(empty(1));

  // Отчет сдан / Оценка / Преподаватель
  const signGrid = [22, 43, 15, 20];
  children.push(tumanTable([
    labelValueRow([{ t: "Отчет сдан «___» __________ 2026 г.", span: 4 }], signGrid),
    labelValueRow([{ t: "", span: 4 }], signGrid),
    labelValueRow([{ t: "Оценка ______________________________", span: 4 }], signGrid),
    labelValueRow([{ t: "", span: 4 }], signGrid),
    labelValueRow([
      { t: "Преподаватель" },
      { t: "Елисеева Н.В., к.т.н., доцент", underline: true },
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

  children.push(empty(4));
  children.push(para([run("МОСКВА 2026")], { alignment: AlignmentType.CENTER }));
  return children;
}

// ── СОДЕРЖАНИЕ ──
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
  body("Лабораторная работа посвящена архитектурному описанию программной системы средствами C4 Model. Требуется научиться отделять систему от её окружения (диаграмма системного контекста, уровень 1) и расчленять её на независимо развёртываемые исполняемые единицы (диаграмма контейнеров, уровень 2). Дополнительная задача — внедрить в повседневную работу принцип Diagrams-as-Code: схемы пишутся текстом в PlantUML с библиотекой C4-PlantUML и автоматически собираются в изображения, как обычный исходный код."),

  h1("2. Постановка задачи"),
  body("Из методического пособия следуют три требования к работе №1:"),
  body("1) системный контекст (System Context Diagram): система показывается цельным «чёрным ящиком»; схема обязана содержать центральный блок, не меньше трёх персон-пользователей и не меньше трёх внешних систем за пределами влияния команды; каждая связь направленная и подписана составом передаваемых данных;"),
  body("2) контейнеры (Container Diagram): центральный блок раскрывается на обособленные, независимо развёртываемые сущности; минимум пять контейнеров (среди них обязательно веб-приложение, мобильное приложение, серверные сервисы, хранилища данных); у каждого контейнера — конкретный технологический стек, у каждой межсервисной связи — сетевой протокол; оформляется таблица контейнеров;"),
  body("3) подготовка отчёта по работе."),

  h1("3. Исходные данные. Общая характеристика системы"),
  body("Моделируется информационная система DemoScope — демографическая аналитическая платформа, которая собирает, наглядно представляет и прогнозирует демографическую картину субъектов Российской Федерации. Система покрывает следующий функционал:"),
  body("– сведение показателей Росстата и ЕМИСС (численность населения, рождаемость, смертность, миграция, заработная плата) в единое хранилище;"),
  body("– интерактивные представления: тепловые карты, 3D-глобус, динамические графики, сравнение территорий;"),
  body("– прогнозирование населения методами Хольта-Уинтерса и ARIMA, когортно-компонентный метод, доверительные интервалы;"),
  body("– «радар проблем» — предиктивные алерты о неблагополучных территориях;"),
  body("– автоматическая выписка аналитических справок с подключением большой языковой модели Google Gemini."),
  body("Проект пребывает на стадии архитектурного проектирования: контуры подсистем зафиксированы целевым образом, наполнение кодом идёт спринтами, поэтому некоторые части (например, мобильный клиент) ещё реализуются. Целевой стек: Python/FastAPI на сервере, React на клиенте; демографические массивы — в озере данных Parquet, служебные данные — в PostgreSQL. Схемы архитектуры поддерживаются как код (нотация C4-PlantUML)."),

  h1("4. Диаграмма системного контекста (Level 1)"),
  body("Диаграмма контекста рисует DemoScope одним сплошным блоком: внутренности скрыты, показаны только граничные взаимодействия. Именно она фиксирует, какие роли пользуются системой и какие внешние сервисы ей необходимы."),
  figure(imgL1),
  caption("Рисунок 1 — Диаграмма системного контекста системы DemoScope"),

  h2("4.1. Спецификация элементов"),
  tableTitle("Таблица 1 — Элементы диаграммы системного контекста"),
  dataTable(
    ["Элемент", "Тип (C4)", "Описание"],
    [
      ["Демограф-аналитик", "Person", "Основной потребитель системы: наблюдает демографические показатели, сравнивает субъекты, строит прогнозы"],
      ["Сотрудник органа государственного управления", "Person", "Берёт «радар проблем» и отчёты за основу решений в социально-демографической сфере"],
      ["Исследователь / студент", "Person", "Изучает динамику населения, выгружает справочные и отчётные материалы"],
      ["DemoScope", "Software System", "Архитектурируемая система: мониторинг, анализ и прогнозирование динамики населения РФ"],
      ["Росстат", "External System", "Первичный источник официальной демографической статистики"],
      ["ЕМИСС", "External System", "Государственный узел публикации статистических данных"],
      ["Google Gemini API", "External System", "Сторонний LLM-сервис (Gemini 2.5 Flash) для аналитических текстов"],
      ["Google Drive", "External System", "Внешнее хранилище исходных CSV-массивов"],
    ],
    [28, 18, 54],
  ),

  h2("4.2. Спецификация связей"),
  tableTitle("Таблица 2 — Связи диаграммы системного контекста"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "DemoScope", "Запросы мониторинга, сравнения территорий и построения прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "DemoScope", "Запросы «радара проблем» и отчётных материалов", "HTTPS"],
      ["Исследователь / студент", "DemoScope", "Просмотр динамики, выгрузка справок и отчётов", "HTTPS"],
      ["DemoScope", "Росстат", "Получение официальной демографической статистики", "HTTPS/REST"],
      ["DemoScope", "ЕМИСС", "Запросы региональной статистики", "HTTPS/REST"],
      ["DemoScope", "Google Gemini API", "Генерация текста заключения через LLM", "HTTPS/JSON"],
      ["DemoScope", "Google Drive", "Загрузка исходных наборов CSV", "HTTPS"],
    ],
    [22, 18, 42, 18],
  ),
  body("Стрелка указывает инициатора и направление потока данных; обратных стрелок на уровне L1 нет. Внутренности системы (контейнеры, компоненты) здесь намеренно не показаны — это прерогатива следующих уровней модели."),

  h1("5. Диаграмма контейнеров (Level 2)"),
  body("На уровне контейнеров DemoScope расчленяется на самостоятельно развёртываемые единицы. Напомним: «контейнер» в C4 — не Docker, а любое приложение или хранилище данных, без которого система не работает как целое. Стек и протоколы зафиксированы по каждому элементу."),
  figure(imgL2),
  caption("Рисунок 2 — Диаграмма контейнеров системы DemoScope"),

  h2("5.1. Таблица описания контейнеров"),
  tableTitle("Таблица 3 — Технологические контейнеры системы"),
  dataTable(
    ["ID", "Контейнер", "Тип", "Тех\u00ADно\u00ADло\u00ADги\u00ADче\u00ADский стек", "Назначение"],
    [
      ["web_app", "Веб-приложение (SPA)", "Container", "React 19 / Vite, Three.js, Recharts, Leaflet", "SPA-клиент: 3D-глобус, тепловые карты, дашборды, графики, сравнение территорий"],
      ["mobile_app", "Мобильное приложение", "Container", "React Native", "Мобильный доступ к мониторингу и прогнозам через единый REST API backend-контейнера"],
      ["api_backend", "Backend API", "Container", "Python 3.10 / FastAPI, Pandas, PyArrow, Statsmodels", "Серверное ядро: REST API мониторинга, сравнения, про\u00ADгно\u00ADзи\u00ADро\u00ADва\u00ADния, отчётов, интеграция с LLM"],
      ["data_pipeline", "Конвейер подготовки данных", "Container", "Python / Pandas, gdown", "Служебный конвейер: выгрузка исходников, очистка, конвертация CSV в Parquet, геомаппинг регионов"],
      ["data_store", "Хранилище де\u00ADмо\u00ADгра\u00ADфи\u00ADче\u00ADских данных", "ContainerDb", "Parquet / PyArrow (озеро данных)", "Массивы Росстата и ЕМИСС: население, рождаемость, смертность, миграция (~6 ГБ)"],
      ["db_service", "Служебная БД", "ContainerDb", "PostgreSQL 16", "Профили и сессии пользователей, история отчётов, кэш прогнозов"],
    ],
    [22, 19, 17, 21, 21],
  ),

  h2("5.2. Спецификация связей"),
  tableTitle("Таблица 4 — Связи диаграммы контейнеров"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "web_app", "Использование дашбордов и прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "web_app", "Получение отчётных материалов и «радара проблем»", "HTTPS"],
      ["web_app", "api_backend", "HTTP-запросы к REST API (/api/population, /api/forecast, /api/report и др.)", "HTTPS/JSON"],
      ["mobile_app", "api_backend", "HTTP-запросы к REST API", "HTTPS/JSON"],
      ["api_backend", "data_store", "Чтение данных и агрегатов", "Файловый I/O (Parquet)"],
      ["api_backend", "db_service", "Обмен служебными данными (аккаунты, сессии, кэш)", "SQL/TCP 5432"],
      ["data_pipeline", "data_store", "Выгрузка подготовленных данных", "Файловый I/O (Parquet)"],
      ["data_pipeline", "Google Drive", "Скачивание исходных CSV", "HTTPS (gdown)"],
      ["api_backend", "Google Gemini API", "Вызов генерации аналитического заключения", "HTTPS/JSON (REST)"],
    ],
    [22, 18, 42, 18],
  ),

  h2("5.3. Обоснование архитектурных решений"),
  body("1. Разнесённые фронтенд и бэкенд. SPA собирается статикой (Vite) и живёт собственным циклом релизов; серверная логика доступна ему только через REST API. Минимальная связанность упрощает независимое обновление клиента и сервера."),
  body("2. Конвейер данных вне процесса API. Подготовка массивов (загрузка, очистка, конвертация) идёт офлайн-утилитами: планировщик запускает их отдельно, и сбой конвейера не валит сервис. Принцип единственной ответственности доведён до уровня развёртывания."),
  body("3. Разделение хранилищ по назначению. Аналитика читает редко, но много — для этого озеро данных на колоночном Parquet с обработкой PyArrow «в памяти»: максимум скорости без администрирования СУБД. Операционным данным (аккаунты, сессии, история отчётов, кэш прогнозов) нужны транзакции — им место в PostgreSQL, доступ по сетевому протоколу SQL/TCP 5432."),
  body("4. Мобильный клиент в общем контуре. Архитектура зафиксирована целевой, реализация распределена по спринтам; мобильное приложение занимает в схеме равное с веб-приложением место и ходит в тот же REST API. React Native достаёт компетенции команды из React 19 и не требует правок других контейнеров."),
  body("5. LLM остаётся снаружи. Тексты заключений выдаёт внешний Google Gemini; собственной модели в DemoScope нет, поэтому её API стоит за границей System_Boundary."),

  h1("6. Глоссарий идентификаторов (Glossary IDs)"),
  body("Перечень единых идентификаторов, которыми помечены элементы на всех диаграммах проекта:"),
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
      ["Система — единый неделимый «чёрный ящик» на L1", "Выполнено: L1 не раскрывает внутренности системы"],
      ["Центральный блок и не менее 3 пользовательских ролей", "Выполнено: 3 персоны на схеме"],
      ["Не менее 3 внешних систем вне зоны контроля разработчиков", "Выполнено: 4 внешние системы за границей контроля"],
      ["Все связи строго направленные, с описанием данных", "Выполнено: 7 направленных связей, все с описанием и протоколом"],
      ["Декомпозиция центрального блока на контейнеры (L2)", "Выполнено: система разделена на 6 контейнеров"],
      ["Не менее 5 контейнеров, включая мобильное и веб-приложение, бэкенд, хранилища", "Выполнено: обязательные типы контейнеров присутствуют: веб- и мобильное приложение, Backend API, конвейер данных, озеро данных Parquet и служебная БД PostgreSQL"],
      ["Технологический стек для каждого контейнера", "Выполнено: стек перечислен в таблице 3"],
      ["Сетевые протоколы для всех межсервисных связей", "Выполнено: протокол проставлен у каждой связи (HTTPS, HTTPS/JSON, SQL/TCP 5432, файловый I/O озера данных)"],
      ["Таблица описания контейнеров", "Выполнено: таблица в п. 5.1"],
      ["Инструмент — PlantUML + C4-PlantUML", "Выполнено: применены макросы Person, System, System_Ext, Container, ContainerDb, Rel, LAYOUT_WITH_LEGEND, LAYOUT_LEFT_RIGHT"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: запрещённые конструкции не использовались"],
    ],
    [55, 45],
  ),

  h1("8. Выводы"),
  body("Отработаны уровни L1 и L2 методологии C4 на архитектуре курсового проекта. Контекстная диаграмма отделила DemoScope от окружения: три пользовательские роли, четыре внешние системы, семь однонаправленных потоков с протоколами. Контейнерная диаграмма разложила систему на шесть независимо развёртываемых единиц с указанием стека и протоколов взаимодействия. Диаграммы существуют в формате Diagrams-as-Code: исходники на C4-PlantUML хранятся в репозитории, PNG собираются автоматически, архитектура версионируется наравне с кодом. Далее следуют Level 3 (компонентная декомпозиция бэкенда) и Level 4 (проектирование классов)."),
];

const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Студент",
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
