// Генератор отчёта по ЛР4 (docx-js) — клон generate_report.js (титульник и оформление — по образцу пользователя)
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб4";
const OUT = path.join(LAB, "отчет_ЛР4.docx");
const LAB_NUM = "4";

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
    labelValueRow([{ t: "«C4 Model. Макроархитектура, динамические процессы и инфраструктура развертывания»", underline: true, align: AlignmentType.CENTER }]),
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

  children.push(...empty(1));
  children.push(para([run("МОСКВА 2026")], { alignment: AlignmentType.CENTER }));
  return children;
}

function buildContents() {
  const items = [
    "1. Цель работы",
    "2. Постановка задачи",
    "3. Исходные данные",
    "4. Диаграмма системного ландшафта (Level 0)",
    "5. Динамическая диаграмма",
    "6. Диаграмма развертывания",
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
const imgL0 = path.join(LAB, "diagrams/png/landscape_l0.png");
const imgDyn = path.join(LAB, "diagrams/png/dynamic_process.png");
const imgDep = path.join(LAB, "diagrams/png/deployment.png");

const bodyChildren = [
  h1("1. Цель работы"),
  body("Получить практические навыки моделирования макроуровня архитектуры и рантайм-поведения системы: научиться отображать интеграцию продукта в корпоративную ИТ-инфраструктуру на диаграмме системного ландшафта (Level 0), фиксировать хронологию сквозного процесса на динамической диаграмме с последовательной нумерацией вызовов, а также проектировать физическую инфраструктуру развертывания контейнеров с указанием узлов, сетевых протоколов, портов и топологии отказоустойчивости."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №4 необходимо:"),
  body("1) разработать диаграмму системного ландшафта (Level 0), отображающую встраивание системы в ИТ-инфраструктуру предметной области с выделением внешних ИТ-границ; отразить взаимодействие целевой системы как минимум с двумя смежными корпоративными системами и одной внешней финансовой системой с обязательным указанием протоколов (gRPC, HTTPS, OAuth2);"),
  body("2) разработать динамическую диаграмму (Dynamic Diagram) для сквозного процесса взаимодействия: модель рантайм-сценария с хронологическим порядком вызовов, зафиксированным последовательной нумерацией;"),
  body("3) разработать диаграмму развертывания для уровня контейнеров: физическая/виртуальная инфраструктура, логические узлы и контейнеры сред выполнения, параметры сетевой безопасности (порты, шифрование TLS, протоколы), топология распределения и отказоустойчивости (репликация БД, обратный прокси);"),
  body("4) оформить отчёт по лабораторной работе."),

  h1("3. Исходные данные"),
  body("Объектом моделирования является система DemoScope (см. лабораторные работы №1–3): платформа демографической аналитики на стеке Python/FastAPI и React, состоящая из шести контейнеров (веб- и мобильное приложение, Backend API, конвейер подготовки данных, озеро данных Parquet, служебная БД PostgreSQL 16). На уровнях L1–L3 описан статический срез архитектуры; в настоящей работе система рассматривается в трёх новых разрезах: макроландшафт ведомственной ИТ-инфраструктуры, рантайм-хронология ключевого процесса и физическое размещение контейнеров."),
  body("Развертывание описывается для целевой конфигурации в облачной инфраструктуре (Yandex Cloud); мобильное приложение включено в контур на общих основаниях (проект находится на этапе архитектурного проектирования, реализация ведётся по спринтам)."),

  h1("4. Диаграмма системного ландшафта (Level 0)"),
  body("Диаграмма ландшафта показывает встраивание DemoScope в ведомственную ИТ-инфраструктуру и взаимодействие с пользовательскими ролями. Граница Enterprise_Boundary очерчивает закрытый корпоративный контур, внутри которого находятся проектируемая система и две смежные корпоративные платформы: ЕСИА (IAM — единая система идентификации) и ведомственная CRM обращений граждан. За пределами контура — три пользовательские роли, внешние источники статистики, LLM-сервис и финансовая система биллинга облачного провайдера."),
  figure(imgL0),
  caption("Рисунок 1 — Диаграмма системного ландшафта DemoScope (Level 0)"),

  h2("4.1. Спецификация элементов"),
  tableTitle("Таблица 1 — Элементы системного ландшафта"),
  dataTable(
    ["Элемент", "Тип", "Расположение", "Назначение"],
    [
      ["Демограф-аналитик", "Person", "Вне контура", "Основной пользователь: мониторит показатели, строит прогнозы и сравнивает территории"],
      ["Сотрудник органа гос. управления", "Person", "Вне контура", "Использует «радар проблем» и аналитические отчёты для принятия управленческих решений"],
      ["Исследователь / студент", "Person", "Вне контура", "Изучает демографическую динамику, выгружает аналитические справки"],
      ["DemoScope", "Software System", "Внутри корпоративного контура", "Проектируемая система мониторинга и прогнозирования"],
      ["ЕСИА (IAM)", "Corporate System", "Внутри корпоративного контура", "Единая система идентификации и аутентификации: SSO, учётные записи"],
      ["Система обращений граждан (CRM)", "Corporate System", "Внутри корпоративного контура", "Регистрация обращений, поручения, контроль сроков"],
      ["Росстат", "External System", "Вне контура", "Официальная демографическая статистика"],
      ["ЕМИСС", "External System", "Вне контура", "Государственная платформа статистических данных"],
      ["Google Gemini API", "External System", "Вне контура", "LLM-сервис аналитических заключений"],
      ["Биллинг облачного провайдера", "External System (финансовая)", "Вне контура", "Учёт потребления и стоимости облачных ресурсов"],
    ],
    [22, 21, 20, 37],
  ),

  h2("4.2. Спецификация связей"),
  tableTitle("Таблица 2 — Связи системного ландшафта"),
  dataTable(
    ["Инициатор", "Получатель", "Назначение", "Протокол"],
    [
      ["Демограф-аналитик", "DemoScope", "Мониторинг показателей и построение прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "DemoScope", "Аналитические отчёты и радар проблем", "HTTPS"],
      ["Исследователь / студент", "DemoScope", "Изучение демографической динамики", "HTTPS"],
      ["DemoScope", "ЕСИА (IAM)", "Сквозная аутентификация пользователей (OAuth2-поток)", "HTTPS / OAuth 2.0"],
      ["CRM", "DemoScope", "Запрос аналитических справок для ответов на обращения", "gRPC / HTTP2"],
      ["DemoScope", "Росстат", "Загрузка демографических показателей", "HTTPS/REST"],
      ["DemoScope", "ЕМИСС", "Загрузка статистических данных по регионам", "HTTPS/REST"],
      ["DemoScope", "Google Gemini API", "Генерация текста аналитических заключений", "HTTPS/JSON"],
      ["DemoScope", "Биллинг облачного провайдера", "Получение счетов и оплата потреблённых ресурсов", "HTTPS/REST"],
    ],
    [17, 22, 39, 22],
  ),
  body("Взаимодействие внутри закрытого контура выполняется по защищённым межсистемным каналам (OAuth 2.0, gRPC). Доступ во внешние сети выполняется только исходящими HTTPS-соединениями из приватного сегмента через контролируемый egress-шлюз с белым списком адресов; публикуемых наружу сервисов у системы нет — входящий трафик ограничен порталом системы (см. п. 6.3)."),

  h1("5. Динамическая диаграмма"),
  body("В качестве сквозного рантайм-процесса выбран ключевой бизнес-сценарий системы — «Построение прогноза численности населения региона». Диаграмма фиксирует хронологический порядок вызовов последовательной нумерацией от 1 до 12: от постановки задачи аналитиком до отображения результата с доверительными интервалами. Вход в контейнер api_backend осуществляется через маршрутизатор api_router, а доступ к озеру данных выполняется строго через компонент «Репозиторий данных» (data_repo) в полном соответствии с трёхслойной архитектурой Backend API (лабораторная работа №2)."),
  figure(imgDyn, 340),
  caption("Рисунок 2 — Динамическая диаграмма: сценарий построения прогноза"),

  h2("5.1. Спецификация шагов"),
  tableTitle("Таблица 3 — Хронология вызовов сквозного сценария"),
  dataTable(
    ["Шаг", "Инициатор", "Получатель", "Действие", "Протокол"],
    [
      ["1", "Демограф-аналитик", "Веб-приложение", "Задаёт параметры прогноза: регион, горизонт, сценарий", "HTTPS"],
      ["2", "Веб-приложение", "API Router", "Запрос прогноза (POST /api/forecast)", "HTTPS/JSON"],
      ["3", "API Router", "Сервис про\u00ADгно\u00ADзи\u00ADро\u00ADвания", "Вызов run_forecast(region, horizon, model_type)", "Python Call"],
      ["4", "Сервис про\u00ADгно\u00ADзи\u00ADро\u00ADвания", "Репозиторий служебной БД", "Поиск результата в кэше прогнозов", "Python Call"],
      ["5", "Репозиторий служебной БД", "Служебная БД", "Запрос SELECT по ключу прогноза", "SQL/TCP 5432"],
      ["6", "Сервис про\u00ADгно\u00ADзи\u00ADро\u00ADвания", "Репозиторий данных", "Запрос исторического ряда", "Python Call"],
      ["7", "Репозиторий данных", "Озеро данных", "Чтение Parquet-фрагментов", "Файловый I/O (Parquet)"],
      ["8", "Сервис про\u00ADгно\u00ADзи\u00ADро\u00ADвания", "Репозиторий служебной БД", "Сохранение рассчитанного прогноза в кэш", "Python Call"],
      ["9", "Репозиторий служебной БД", "Служебная БД", "Запрос INSERT в кэш прогнозов", "SQL/TCP 5432"],
      ["10", "Сервис про\u00ADгно\u00ADзи\u00ADро\u00ADвания", "API Router", "Возврат структуры ForecastResult", "Python Call"],
      ["11", "API Router", "Веб-приложение", "HTTP 200 OK: прогноз с доверительными интервалами", "HTTPS/JSON"],
      ["12", "Веб-приложение", "Демограф-аналитик", "Отображение графика прогноза и доверительных интервалов", "HTTPS"],
    ],
    [9, 21, 20, 33, 17],
  ),
  body("Шаги 4–5 реализуют проверку кэша до тяжёлого расчёта: при попадании в кэш повторный запрос того же прогноза обслуживается из СУБД, а шаги 6–9 (чтение озера данных и сохранение результата) пропускаются — это снижает время отклика и нагрузку на хранилище."),

  h1("6. Диаграмма развертывания"),
  body("Диаграмма развертывания отображает физическую инфраструктуру, на которую разворачиваются контейнеры системы: клиентские узлы (рабочая станция, смартфон), облачная инфраструктура Yandex Cloud с узлами веб-сервисов, приложений (Docker), пакетной обработки, озера данных и кластера PostgreSQL, а также внешние системы."),
  figure(imgDep),
  caption("Рисунок 3 — Диаграмма развертывания DemoScope"),

  h2("6.1. Узлы и размещённые контейнеры"),
  tableTitle("Таблица 4 — Узлы развертывания"),
  dataTable(
    ["Узел", "Тип", "Контейнеры"],
    [
      ["Рабочая станция аналитика", "Windows / macOS", "Веб-приложение (SPA React 19, браузер)"],
      ["Смартфон пользователя", "Android / iOS", "Мобильное приложение (React Native)"],
      ["Сервер веб-сервисов", "Ubuntu Linux VM (публичный IP)", "Reverse Proxy (Nginx): маршрутизация, SSL-терминация, статика SPA"],
      ["Сервер приложений", "Ubuntu Linux (Docker Host)", "Docker Engine 24 → Backend API (Python 3.10 / FastAPI)"],
      ["Сервер пакетной обработки", "Ubuntu Linux (Cron)", "Конвейер подготовки данных (Python / Pandas)"],
      ["Сервер озера данных", "Ubuntu Linux VM (NFS-том)", "Озеро данных (Parquet / PyArrow, ~6 ГБ)"],
      ["Сервер БД (Мастер)", "Managed PostgreSQL VM", "Служебная БД (PostgreSQL 16) — все транзакции записи"],
      ["Сервер БД (Реплика)", "Managed PostgreSQL VM", "Реплика БД (PostgreSQL 16) — горячий резерв"],
    ],
    [24, 26, 50],
  ),

  h2("6.2. Сетевые взаимодействия"),
  tableTitle("Таблица 5 — Сетевые интерфейсы и протоколы"),
  dataTable(
    ["Исходный элемент", "Целевой элемент", "Описание взаимодействия", "Протокол и порт", "Сетевая среда"],
    [
      ["Веб-приложение (SPA)", "Reverse Proxy", "Запросы к системе", "HTTPS/TLS, порт 443", "Открытая сеть (Internet)"],
      ["Мобильное приложение", "Reverse Proxy", "Запросы к API", "HTTPS/TLS, порт 443", "Открытая сеть (Internet)"],
      ["Reverse Proxy", "Backend API", "Проксирование запросов", "HTTP/JSON, порт 8000", "Внутренняя сеть ЦОД"],
      ["Backend API", "Служебная БД (Мастер)", "Запись и чтение служебных данных", "SQL/TCP, порт 5432", "Сеть БД (выделенная подсеть)"],
      ["Служебная БД (Мастер)", "Реплика БД", "Потоковая репликация", "WAL/TCP, порт 5432", "Изолированная сеть репликации"],
      ["Backend API", "Служебная БД (Реплика)", "Аналитические выборки (чтение)", "SQL/TCP, порт 5432", "Сеть БД (выделенная подсеть)"],
      ["Backend API", "Озеро данных", "Чтение Parquet-фрагментов", "Файловый I/O (NFS), порт 2049", "Внутренняя сеть ЦОД"],
      ["Конвейер подготовки данных", "Озеро данных", "Запись подготовленных данных", "Файловый I/O (NFS), порт 2049", "Внутренняя сеть ЦОД"],
      ["Конвейер подготовки данных", "Росстат", "Выгрузка официальных показателей", "HTTPS/REST, порт 443", "Открытая сеть (Internet)"],
      ["Конвейер подготовки данных", "ЕМИСС", "Выгрузка региональных статистических витрин", "HTTPS/REST, порт 443", "Открытая сеть (Internet)"],
      ["Конвейер подготовки данных", "Google Drive", "Скачивание исходных массивов (CSV)", "HTTPS (gdown), порт 443", "Открытая сеть (Internet)"],
      ["Backend API", "Google Gemini API", "Генерация аналитических заключений", "HTTPS/JSON, порт 443", "Открытая сеть (Internet)"],
      ["Backend API", "ЕСИА (IAM)", "Проверка токенов OAuth 2.0", "HTTPS, порт 443", "Открытая сеть (Internet)"],
    ],
    [17, 15, 26, 21, 21],
  ),

  h2("6.3. Политика безопасности периметра"),
  body("1. Публичный контур. Прямой доступ из интернета разрешён только к порту 443 (Nginx) на сервере веб-сервисов с публичным IP; трафик шифруется TLS, SSL-терминация выполняется на обратном прокси."),
  body("2. Приватный контур. Сервер приложений, сервер пакетной обработки, узел озера данных и серверы БД закрыты от внешнего мира межсетевыми экранами и доступны только по внутренним сетевым интерфейсам из доверенных сегментов ЦОД (VPC)."),
  body("3. Сегментация сетей БД. Транзакции записи принимает только мастер; репликация WAL выполняется в изолированной сети репликации, тяжёлые аналитические выборки разгружаются на реплику."),
  body("4. Секреты и доступы. Учётные данные для внешних интеграций (Росстат, ЕМИСС, Gemini, Google Drive) передаются контейнерам через переменные окружения оркестратора; ключи OAuth2 для ЕСИА хранятся в защищённом хранилище секретов; доступ к NFS-озеру данных ограничен экспортом томов только подсетям сервера приложений и сервера пакетной обработки."),

  h1("7. Глоссарий идентификаторов (Glossary IDs)"),
  body("Идентификаторы ландшафта и инфраструктуры развивают глоссарий ЛР1: контейнеры системы сохраняют канонические ID на всех диаграммах, включая развертывание:"),
  tableTitle("Таблица 6 — Глоссарий идентификаторов ландшафта и инфраструктуры"),
  dataTable(
    ["ID", "Элемент", "Тип"],
    [
      ["esia", "ЕСИА (IAM)", "Смежная корпоративная система (L0 и развертывание)"],
      ["crm", "Система обращений граждан (CRM)", "Смежная корпоративная система (L0)"],
      ["billing", "Биллинг облачного провайдера", "Внешняя финансовая система (L0)"],
      ["web_app, mobile_app, api_backend, data_pipeline, data_store", "Контейнеры системы, развёрнутые на узлах (браузер SPA, смартфон, Docker-хост, Cron, сервер озера данных)", "Контейнеры (глоссарий ЛР1)"],
      ["nginx", "Reverse Proxy (Nginx)", "Контейнер среды выполнения (развертывание)"],
      ["db_master", "Служебная БД (Мастер)", "ContainerDb на узле «Сервер БД (Мастер)»"],
      ["db_replica", "Реплика БД (Standby)", "ContainerDb на узле «Сервер БД (Реплика)»"],
    ],
    [22, 42, 36],
  ),

  h1("8. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 7 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Ландшафт Level 0: встраивание в ИТ-инфраструктуру с выделением внешних границ", "Выполнено: граница Enterprise_Boundary «Ведомственная ИТ-инфраструктура», системы и пользователи внутри и вне контура"],
      ["Взаимодействие не менее чем с 2 смежными корпоративными системами", "Выполнено: ЕСИА (IAM) и ведомственная CRM"],
      ["Взаимодействие с 1 внешней финансовой системой", "Выполнено: биллинг облачного провайдера"],
      ["Обязательные протоколы (gRPC, HTTPS, OAuth2)", "Выполнено: gRPC/HTTP2 (CRM), HTTPS/OAuth 2.0 (ЕСИА), HTTPS/REST"],
      ["Динамическая диаграмма сквозного процесса", "Выполнено: сценарий построения прогноза, 12 шагов"],
      ["Хронологический порядок вызовов с последовательной нумерацией", "Выполнено: нумерация 1–12 на стрелках и в таблице шагов"],
      ["Диаграмма развертывания: физическая/виртуальная инфраструктура", "Выполнено: 2 клиентских узла + 6 VM-узлов ЦОД (веб-сервисы, Docker-хост, пакетная обработка, озеро данных, мастер и реплика БД)"],
      ["Логические узлы и контейнеры сред выполнения", "Выполнено: Nginx, Docker Engine, контейнеры Backend API и конвейера; идентификаторы контейнеров — строго по глоссарию ЛР1"],
      ["Параметры сетевой безопасности (порты, TLS, протоколы)", "Выполнено: порты 443/8000/5432/2049, TLS, NFS, WAL; политика периметра в п. 6.3"],
      ["Топология распределения и отказоустойчивости", "Выполнено: репликация БД Master–Replica (WAL), Nginx Reverse Proxy"],
    ],
    [55, 45],
  ),

  h1("9. Выводы"),
  body("В ходе лабораторной работы архитектурное описание DemoScope дополнено тремя представлениями за пределами классического «зумирования». Диаграмма системного ландшафта (Level 0) показала встраивание системы в ведомственную ИТ-инфраструктуру: две смежные корпоративные платформы (ЕСИА, CRM) и внешняя финансовая система (биллинг облачного провайдера) интегрированы по протоколам OAuth 2.0, gRPC и HTTPS. Динамическая диаграмма зафиксировала хронологию сквозного сценария построения прогноза из двенадцати нумерованных шагов (с прохождением через маршрутизатор api_router, проверкой кэша прогнозов и чтением озера данных через data_repo). Диаграмма развертывания отобразила физическое размещение контейнеров на узлах облачной инфраструктуры с указанием портов (443, 8000, 5432, 2049), шифрования TLS, доступа по NFS, сетевой сегментации и топологии отказоустойчивости (репликация PostgreSQL Master–Replica, Nginx Reverse Proxy). Архитектурное описание системы теперь покрывает все уровни модели C4 — от ландшафта до кода — и подготовлено к интеграции в единый репозиторий с автоматической сборкой (лабораторная работа №5)."),
];
const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Бектеев Георгий",
  title: "Отчёт по лабораторной работе №4 — C4 Model. Макроархитектура, динамика, развертывание",
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
