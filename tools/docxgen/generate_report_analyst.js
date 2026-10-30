// Генератор отчёта по ЛР1 — вариант для участника с ролью аналитик (лаб1_аналитик).
// Титульник и оформление — те же, что в базовом отчёте (образец СТАНКИН);
// основной текст переписан другими формулировками, диаграммы — из лаб1_аналитик/diagrams.
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб1_аналитик";
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
  body("В рамках лабораторной работы осваивается методология C4 Model применительно к архитектуре собственного курсового проекта. Ставятся две практические задачи: закрепить внешние границы системы на диаграмме системного контекста (уровень 1) и разделить систему на независимо развёртываемые технологические контейнеры на диаграмме контейнеров (уровень 2). Попутно отрабатывается инженерная практика Diagrams-as-Code: архитектурные схемы хранятся в текстовом виде (PlantUML с библиотекой C4-PlantUML) и рендерятся в PNG автоматически, что позволяет держать их в системе контроля версий рядом с программным кодом."),

  h1("2. Постановка задачи"),
  body("Лабораторная работа №1 по методическому пособию состоит из трёх частей:"),
  body("1) контекстная диаграмма (System Context Diagram). Проектируемая система изображается одним неделимым «чёрным ящиком»; на схеме присутствуют центральный блок, минимум три пользовательские роли и минимум три внешние программные или аппаратные системы, не контролируемые командой; все связи задаются направленными и снабжаются описанием передаваемых данных;"),
  body("2) контейнерная диаграмма (Container Diagram). «Чёрный ящик» декомпозируется на обособленные, независимо развёртываемые исполняемые сущности; требуется выделить не меньше пяти технологических контейнеров (обязательны веб-приложение, мобильное приложение, серверные сервисы и хранилища данных), определить стек каждого контейнера и указать сетевые протоколы всех межсервисных взаимодействий; результаты сводятся в таблицу описания контейнеров;"),
  body("3) оформление отчёта по лабораторной работе."),

  h1("3. Исходные данные. Общая характеристика системы"),
  body("Объект моделирования — информационная система DemoScope, платформа демографической аналитики для мониторинга, исследования и прогнозирования демографической обстановки в регионах Российской Федерации. Ключевые функции системы:"),
  body("– консолидация показателей Росстата и ЕМИСС (население, рождаемость, смертность, миграция, заработная плата) в едином хранилище;"),
  body("– наглядная визуализация: тепловые карты, 3D-глобус, динамические графики, режим сопоставления территорий;"),
  body("– прогнозная аналитика численности населения: модели Хольта-Уинтерса и ARIMA, когортно-компонентный метод, доверительные интервалы;"),
  body("– «радар проблем» — автоматическое выявление неблагополучных территорий по предиктивным алертам;"),
  body("– генерация аналитических справок, включая автоматический текст от большой языковой модели (Google Gemini)."),
  body("Проект находится на этапе архитектурного проектирования: состав подсистем определяется целевым образом, реализация идёт поэтапно (спринты), поэтому отдельные части — например, мобильное приложение — пока в разработке. Целевая архитектура: Python/FastAPI (сервер) и React (клиент); демографические массивы размещаются в озере данных формата Parquet, служебные данные — в СУБД PostgreSQL. Архитектурные схемы ведутся как код в нотации C4-PlantUML."),

  h1("4. Диаграмма системного контекста (Level 1)"),
  body("Уровень системного контекста отвечает на два вопроса: кому система полезна и чем она обменивается с внешним миром. DemoScope изображён одним цельным блоком — «чёрным ящиком», без деталей внутреннего устройства; вокруг него показаны пользовательские роли и внешние системы со строго направленными потоками данных."),
  figure(imgL1),
  caption("Рисунок 1 — Диаграмма системного контекста системы DemoScope"),

  h2("4.1. Спецификация элементов"),
  tableTitle("Таблица 1 — Элементы диаграммы системного контекста"),
  dataTable(
    ["Элемент", "Тип (C4)", "Описание"],
    [
      ["Демограф-аналитик", "Person", "Главный пользователь: следит за демографическими показателями, сопоставляет территории, заказывает прогнозы"],
      ["Сотрудник органа государственного управления", "Person", "Работает с «радаром проблем» и отчётными материалами при подготовке мер социально-демографической политики"],
      ["Исследователь / студент", "Person", "Анализирует демографические процессы, забирает справки и готовые отчёты"],
      ["DemoScope", "Software System", "Моделируемая платформа: мониторинг, анализ и прогнозирование динамики населения РФ"],
      ["Росстат", "External System", "Первоисточник официальной статистики по населению"],
      ["ЕМИСС", "External System", "Государственный портал статистических показателей"],
      ["Google Gemini API", "External System", "Облачный LLM (Gemini 2.5 Flash), готовящий тексты аналитических заключений"],
      ["Google Drive", "External System", "Облачное хранилище исходных наборов данных (CSV) для конвейера загрузки"],
    ],
    [28, 18, 54],
  ),

  h2("4.2. Спецификация связей"),
  tableTitle("Таблица 2 — Связи диаграммы системного контекста"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "DemoScope", "Просмотр показателей, сопоставление территорий, расчёт прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "DemoScope", "Работа с «радаром проблем», получение аналитических отчётов", "HTTPS"],
      ["Исследователь / студент", "DemoScope", "Просмотр динамики, скачивание справок и отчётов", "HTTPS"],
      ["DemoScope", "Росстат", "Сбор официальных демографических показателей", "HTTPS/REST"],
      ["DemoScope", "ЕМИСС", "Импорт региональной статистики", "HTTPS/REST"],
      ["DemoScope", "Google Gemini API", "Подготовка текста аналитического заключения средствами LLM", "HTTPS/JSON"],
      ["DemoScope", "Google Drive", "Получение исходных наборов данных (CSV)", "HTTPS"],
    ],
    [22, 18, 42, 18],
  ),
  body("Направление каждой стрелки задаёт инициатора обмена; двусторонних связей на уровне контекста нет. Внутреннее устройство системы — контейнеры, компоненты — на этой диаграмме не раскрывается по определению уровня L1."),

  h1("5. Диаграмма контейнеров (Level 2)"),
  body("Контейнерная диаграмма показывает, из каких независимо развёртываемых частей собран «чёрный ящик» L1. Важно, что «контейнер» здесь — термин C4: это исполняемое приложение или хранилище данных, а не единица Docker. Для каждой части указан технологический стек, для каждой связи — протокол."),
  figure(imgL2),
  caption("Рисунок 2 — Диаграмма контейнеров системы DemoScope"),

  h2("5.1. Таблица описания контейнеров"),
  tableTitle("Таблица 3 — Технологические контейнеры системы"),
  dataTable(
    ["ID", "Контейнер", "Тип", "Тех\u00ADно\u00ADло\u00ADги\u00ADче\u00ADский стек", "Назначение"],
    [
      ["web_app", "Веб-приложение (SPA)", "Container", "React 19 / Vite, Three.js, Recharts, Leaflet", "Экраны пользователя: 3D-глобус, тепловые карты, дашборды, графики, режим сравнения территорий"],
      ["mobile_app", "Мобильное приложение", "Container", "React Native", "Смартфон-доступ к мониторингу и прогнозам через единый REST API backend-контейнера"],
      ["api_backend", "Backend API", "Container", "Python 3.10 / FastAPI, Pandas, PyArrow, Statsmodels", "Серверная логика: REST API мониторинга, сопоставления территорий, про\u00ADгно\u00ADзи\u00ADро\u00ADва\u00ADния, отчётов и LLM-анализа"],
      ["data_pipeline", "Конвейер подготовки данных", "Container", "Python / Pandas, gdown", "Сервисные утилиты: выгрузка исходников, очистка, перек\u00ADо\u00ADди\u00ADро\u00ADва\u00ADние CSV в Parquet, геомаппинг регионов"],
      ["data_store", "Хранилище де\u00ADмо\u00ADгра\u00ADфи\u00ADче\u00ADских данных", "ContainerDb", "Parquet / PyArrow (озеро данных)", "Массивы показателей Росстата и ЕМИСС: население, рождаемость, смертность, миграция (~6 ГБ)"],
      ["db_service", "Служебная БД", "ContainerDb", "PostgreSQL 16", "Учётные записи и сессии пользователей, история отчётов, кэш прогнозов"],
    ],
    [22, 19, 17, 21, 21],
  ),

  h2("5.2. Спецификация связей"),
  tableTitle("Таблица 4 — Связи диаграммы контейнеров"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "web_app", "Просмотр дашбордов и прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "web_app", "Запрос отчётов и «радара проблем»", "HTTPS"],
      ["web_app", "api_backend", "Обращения к REST API (/api/population, /api/forecast, /api/report и др.)", "HTTPS/JSON"],
      ["mobile_app", "api_backend", "Обращения к REST API", "HTTPS/JSON"],
      ["api_backend", "data_store", "Выборка показателей и агрегатов", "Файловый I/O (Parquet)"],
      ["api_backend", "db_service", "Доступ к служебным данным (аккаунты, сессии, кэш)", "SQL/TCP 5432"],
      ["data_pipeline", "data_store", "Сохранение подготовленных данных", "Файловый I/O (Parquet)"],
      ["data_pipeline", "Google Drive", "Получение исходных наборов (CSV)", "HTTPS (gdown)"],
      ["api_backend", "Google Gemini API", "Подготовка текста заключения через LLM", "HTTPS/JSON (REST)"],
    ],
    [22, 18, 42, 18],
  ),

  h2("5.3. Обоснование архитектурных решений"),
  body("1. Клиент отделён от сервера. Веб-приложение живёт в собственном контейнере и доставляется как статическая сборка Vite; единственный канал к серверной логике — REST API. Клиент и сервер можно обновлять независимо друг от друга."),
  body("2. Подготовка данных не нагружает API. Загрузка, очистка и конвертация массивов вынесены в отдельные офлайн-утилиты: конвейер стартует по расписанию сам по себе и не влияет на работающий сервис. Каждый компонент отвечает за одно дело — уже на уровне развёртывания."),
  body("3. Двухуровневая организация данных. Редкие чтения крупных неизменяемых массивов обслуживает озеро данных на колоночном Parquet под управлением PyArrow «в памяти» — высокая скорость аналитических выборок без администрирования тяжёлой СУБД. Служебные же данные (учётные записи, сессии, история отчётов, кэш прогнозов) живут в PostgreSQL: транзакционность гарантирована СУБД, а связь идёт по стандартному сетевому протоколу SQL/TCP 5432."),
  body("4. Мобильное приложение в контуре системы. Проект проходит стадию архитектурного проектирования: контейнеры задаются целевыми и наполняются реализацией по спринтам. Мобильный клиент стоит в архитектуре наравне с веб-приложением и делит с ним единый REST API; стек React Native опирается на накопленные компетенции команды (React 19) и не требует правок остальных контейнеров."),
  body("5. Генерация текстов — снаружи. Аналитические заключения пишет внешний сервис Google Gemini; собственной модели в системе нет, поэтому её API вынесен за границу System_Boundary."),

  h1("6. Глоссарий идентификаторов (Glossary IDs)"),
  body("Список сквозных идентификаторов, применяемых во всех схемах проекта:"),
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
      ["Система — единый неделимый «чёрный ящик» на L1", "Выполнено: внутренности системы на L1 не показываются"],
      ["Центральный блок и не менее 3 пользовательских ролей", "Выполнено: 3 персоны"],
      ["Не менее 3 внешних систем вне зоны контроля разработчиков", "Выполнено: 4 внешние системы"],
      ["Все связи строго направленные, с описанием данных", "Выполнено: 7 направленных связей, у каждой описание и протокол"],
      ["Декомпозиция центрального блока на контейнеры (L2)", "Выполнено: 6 контейнеров"],
      ["Не менее 5 контейнеров, включая мобильное и веб-приложение, бэкенд, хранилища", "Выполнено: все обязательные категории на месте: веб- и мобильное приложение, Backend API, конвейер данных, озеро данных Parquet и служебная БД PostgreSQL"],
      ["Технологический стек для каждого контейнера", "Выполнено: указан в таблице 3"],
      ["Сетевые протоколы для всех межсервисных связей", "Выполнено: протокол есть у каждой связи (HTTPS, HTTPS/JSON, SQL/TCP 5432, файловый I/O озера данных)"],
      ["Таблица описания контейнеров", "Выполнено: п. 5.1"],
      ["Инструмент — PlantUML + C4-PlantUML", "Выполнено: задействованы макросы Person, System, System_Ext, Container, ContainerDb, Rel, LAYOUT_WITH_LEGEND, LAYOUT_LEFT_RIGHT"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: нарушений нет"],
    ],
    [55, 45],
  ),

  h1("8. Выводы"),
  body("Лабораторная работа дала практику применения C4 Model к курсовому проекту. На уровне контекста зафиксированы внешние границы DemoScope: три роли пользователей, четыре внешние системы, семь однонаправленных потоков данных с протоколами. На уровне контейнеров система разложена на шесть независимо развёртываемых частей, для каждой определён стек, для каждой межсервисной связи — протокол. Схемы оформлены в Diagrams-as-Code: их исходники лежат в репозитории рядом с кодом и автоматически превращаются в PNG, поэтому архитектура версионируется вместе с разработкой. Построенные уровни станут отправной точкой для Level 3 (компоненты бэкенда) и Level 4 (классы)."),
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
