// Генератор отчёта по ЛР1 — вариант для участника с ролью tech-doc (лаб1_другая).
// Титульник и оформление — те же, что в базовом отчёте (образец СТАНКИН);
// основной текст переписан другими формулировками, диаграммы — из лаб1_другая/diagrams.
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб1_другая";
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
  body("Цель работы — освоить методологию C4 Model как инструмент архитектурного проектирования программных систем. В ходе выполнения работы необходимо научиться фиксировать внешние границы проектируемой системы на диаграмме системного контекста (уровень 1) и декомпозировать систему на независимо развёртываемые технологические контейнеры на диаграмме контейнеров (уровень 2). Отдельное внимание уделяется практике Diagrams-as-Code: архитектурные схемы описываются в виде текста на языке PlantUML с подключением библиотеки C4-PlantUML и автоматически преобразуются в графические изображения, что позволяет версионировать архитектуру совместно с кодом проекта."),

  h1("2. Постановка задачи"),
  body("Согласно методическому пособию, в рамках лабораторной работы №1 требуется:"),
  body("1) построить диаграмму системного контекста (System Context Diagram): проектируемая система изображается единым неделимым «чёрным ящиком»; в состав диаграммы входят центральный блок, не менее трёх ролей пользователей (персоны) и не менее трёх внешних программных или аппаратных систем, находящихся вне зоны контроля команды разработки; каждая связь должна быть направленной и сопровождаться описанием состава передаваемых данных;"),
  body("2) построить диаграмму контейнеров (Container Diagram): раскрыть содержимое центрального блока, декомпозировав его на обособленные, независимо развёртываемые исполняемые сущности; выделить не менее пяти технологических контейнеров, среди которых должны быть веб-приложение, мобильное приложение, серверные сервисы и хранилища данных; для каждого контейнера указать технологический стек, для каждой межсервисной связи — сетевой протокол; дополнить диаграмму таблицей описания контейнеров;"),
  body("3) оформить отчёт по результатам работы."),

  h1("3. Исходные данные. Общая характеристика системы"),
  body("Архитектурному моделированию подвергается информационная система DemoScope — платформа демографической аналитики, обеспечивающая мониторинг, анализ и прогнозирование демографической ситуации в субъектах Российской Федерации. Назначение системы раскрывается через следующий круг задач:"),
  body("– сбор в едином хранилище демографических и социально-экономических показателей, публикуемых Росстатом и ЕМИСС: численность населения, рождаемость, смертность, миграция, уровень заработной платы;"),
  body("– интерактивное представление данных: тепловые карты, трёхмерный глобус, динамические графики, сопоставление территорий между собой;"),
  body("– прогнозирование численности населения методами Хольта-Уинтерса, ARIMA и когортно-компонентным методом с построением доверительных интервалов;"),
  body("– формирование «радара проблем» — выявления неблагополучных территорий на базе предиктивных алертов;"),
  body("– автоматизированная подготовка аналитических справок, в том числе с привлечением большой языковой модели Google Gemini."),
  body("Серверная часть системы реализована на Python/FastAPI, клиентская — на React; данные хранятся в колоночном формате Parquet. Архитектурное описание ведётся в подходе Diagrams-as-Code средствами нотации C4-PlantUML."),

  h1("4. Диаграмма системного контекста (Level 1)"),
  body("На диаграмме системного контекста проектируемая система показана в окружении внешних участников с максимально высокого уровня абстракции: DemoScope условно «запечатан» в чёрный ящик, и его внутреннее устройство на схеме не раскрывается. Диаграмма фиксирует, какую ценность система несёт каждой пользовательской роли и через какие интерфейсы она обменивается данными с внешними системами."),
  figure(imgL1),
  caption("Рисунок 1 — Диаграмма системного контекста системы DemoScope"),

  h2("4.1. Спецификация элементов"),
  tableTitle("Таблица 1 — Элементы диаграммы системного контекста"),
  dataTable(
    ["Элемент", "Тип (C4)", "Описание"],
    [
      ["Демограф-аналитик", "Person", "Ключевая роль: отслеживает демографические показатели, сопоставляет регионы, запрашивает прогнозы"],
      ["Сотрудник органа государственного управления", "Person", "Опирается на «радар проблем» и аналитические отчёты при выработке решений в социально-демографической политике"],
      ["Исследователь / студент", "Person", "Изучает демографические процессы, выгружает справочные и отчётные материалы"],
      ["DemoScope", "Software System", "Моделируемая система: мониторинг, анализ и прогнозирование динамики населения РФ"],
      ["Росстат", "External System", "Поставщик официальной демографической статистики"],
      ["ЕМИСС", "External System", "Государственная платформа публикации статистических данных"],
      ["Google Gemini API", "External System", "Внешний LLM-сервис, формирующий тексты аналитических заключений (Gemini 2.5 Flash)"],
      ["Google Drive", "External System", "Репозиторий исходных массивов данных (CSV), откуда их забирает конвейер загрузки"],
    ],
    [28, 18, 54],
  ),

  h2("4.2. Спецификация связей"),
  tableTitle("Таблица 2 — Связи диаграммы системного контекста"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "DemoScope", "Запросы показателей, сравнения регионов и расчёта прогнозов", "HTTPS"],
      ["Сотрудник органа гос. управления", "DemoScope", "Обращения к «радару проблем» и аналитическим отчётам", "HTTPS"],
      ["Исследователь / студент", "DemoScope", "Запросы справок, выгрузка отчётов", "HTTPS"],
      ["DemoScope", "Росстат", "Загрузка официальных демографических показателей", "HTTPS/REST"],
      ["DemoScope", "ЕМИСС", "Получение статистики по субъектам РФ", "HTTPS/REST"],
      ["DemoScope", "Google Gemini API", "Формирование текста аналитического заключения", "HTTPS/JSON"],
      ["DemoScope", "Google Drive", "Скачивание исходных CSV-массивов данных", "HTTPS"],
    ],
    [22, 18, 42, 18],
  ),
  body("Каждая связь на диаграмме однонаправленная: стрелка показывает, кто инициировал обмен и в какую сторону движутся данные. Уровень L1 сознательно не раскрывает внутренности системы — ни контейнеры, ни компоненты здесь не изображаются."),

  h1("5. Диаграмма контейнеров (Level 2)"),
  body("Диаграмма контейнеров «открывает» чёрный ящик уровня L1: центральный блок разбит на независимо развёртываемые исполняемые сущности. Подчеркнём, что контейнер в терминах C4 — это не Docker-единица, а любое приложение или хранилище данных, необходимое для работы системы как целого. Технологический стек зафиксирован для каждого контейнера, а для каждой межсервисной связи указан применяемый протокол."),
  figure(imgL2),
  caption("Рисунок 2 — Диаграмма контейнеров системы DemoScope"),

  h2("5.1. Таблица описания контейнеров"),
  tableTitle("Таблица 3 — Технологические контейнеры системы"),
  dataTable(
    ["ID", "Контейнер", "Тип", "Тех\u00ADно\u00ADло\u00ADги\u00ADче\u00ADский стек", "Назначение"],
    [
      ["web_app", "Веб-приложение (SPA)", "Container", "React 19 / Vite, Three.js, Recharts, Leaflet", "Клиентский интерфейс: 3D-глобус, тепловые карты, дашборды, графики, сопоставление территорий"],
      ["mobile_app", "Мобильное приложение (планируемое)", "Container", "React Native", "Мобильный клиент мониторинга и прогнозов. Контейнер уровня To-Be: принадлежит целевой архитектуре, находится в разработке"],
      ["api_backend", "Backend API", "Container", "Python 3.10 / FastAPI, Pandas, PyArrow, Statsmodels", "Ядро бизнес-логики: REST API мониторинга, сопоставления территорий, про\u00ADгно\u00ADзи\u00ADро\u00ADва\u00ADния, отчётности и LLM-анализа"],
      ["data_pipeline", "Конвейер подготовки данных", "Container", "Python / Pandas, gdown", "Служебный инструментарий: выгрузка исходных массивов, очистка, пересчёт CSV в Parquet, геомаппинг регионов"],
      ["data_store", "Хранилище де\u00ADмо\u00ADгра\u00ADфи\u00ADче\u00ADских данных", "ContainerDb", "Parquet / PyArrow (файловое хранилище)", "Хранение показателей Росстата и ЕМИСС: население, рождаемость, смертность, миграция (~6 ГБ)"],
    ],
    [22, 19, 17, 21, 21],
  ),

  h2("5.2. Спецификация связей"),
  tableTitle("Таблица 4 — Связи диаграммы контейнеров"),
  dataTable(
    ["Инициатор", "Получатель", "Передаваемые данные / бизнес-действие", "Протокол"],
    [
      ["Демограф-аналитик", "web_app", "Работа с дашбордами и прогнозами", "HTTPS"],
      ["Сотрудник органа гос. управления", "web_app", "Получение отчётов и «радара проблем»", "HTTPS"],
      ["web_app", "api_backend", "Вызовы REST API (/api/population, /api/forecast, /api/report и др.)", "HTTPS/JSON"],
      ["mobile_app", "api_backend", "Вызовы REST API (To-Be)", "HTTPS/JSON"],
      ["api_backend", "data_store", "Чтение показателей и агрегатов", "Файловый I/O (Parquet)"],
      ["data_pipeline", "data_store", "Запись подготовленных массивов", "Файловый I/O (Parquet)"],
      ["data_pipeline", "Google Drive", "Загрузка исходных массивов (CSV)", "HTTPS (gdown)"],
      ["api_backend", "Google Gemini API", "Запрос текста аналитического заключения", "HTTPS/JSON (REST)"],
    ],
    [22, 18, 42, 18],
  ),

  h2("5.3. Обоснование архитектурных решений"),
  body("1. Изоляция клиентского слоя от серверного. Веб-приложение оформлено самостоятельным контейнером: его статическая сборка (Vite) развёртывается независимо, а взаимодействие с сервером идёт исключительно через REST API. Такой разрыв связанности позволяет развивать интерфейс и бэкенд в разных темпах и разными специалистами."),
  body("2. Конвейер данных как отдельный контейнер. Загрузка, очистка и конвертация исходных массивов реализованы офлайн-утилитами, не входящими в процесс API-сервиса. Конвейер можно запускать по расписанию вне зависимости от работоспособности бэкенда — это прямое следование принципу единственной ответственности, вынесенное на уровень развёртывания."),
  body("3. Parquet вместо классической СУБД. Нагрузка системы характеризуется редким чтением больших неизменяемых массивов, что идеально ложится на колоночный формат Parquet: PyArrow обрабатывает файлы «в памяти» без затрат на администрирование СУБД, сохраняя высокую скорость аналитических выборок."),
  body("4. Мобильное приложение как контейнер To-Be. Методика C4 разрешает документировать не только текущую (as-is), но и целевую (to-be) архитектуру. Мобильный клиент присутствует на диаграмме как планируемый контейнер с пометкой [To-Be]; стек React Native выбран из-за переиспользования наработанной командой экспертизы в React 19, а подключение через существующий REST API не требует перестройки остальных контейнеров."),
  body("5. LLM-генерация вынесена за границу системы. Тексты аналитических заключений порождает внешний сервис Google Gemini: собственная модель внутри DemoScope отсутствует, поэтому API размещён за пределами System_Boundary."),

  h1("6. Глоссарий идентификаторов (Glossary IDs)"),
  body("Сквозные текстовые идентификаторы, единые для всех диаграмм проекта:"),
  tableTitle("Таблица 5 — Глоссарий идентификаторов"),
  dataTable(
    ["ID", "Элемент", "Тип"],
    [
      ["analyst", "Демограф-аналитик", "Пользователь"],
      ["official", "Сотрудник органа гос. управления", "Пользователь"],
      ["researcher", "Исследователь / студент", "Пользователь"],
      ["demoscope", "DemoScope (граница системы)", "Software System"],
      ["web_app", "Веб-приложение (SPA)", "Контейнер"],
      ["mobile_app", "Мобильное приложение (To-Be)", "Контейнер"],
      ["api_backend", "Backend API", "Контейнер"],
      ["data_pipeline", "Конвейер подготовки данных", "Контейнер"],
      ["data_store", "Хранилище демографических данных", "Контейнер (БД)"],
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
      ["Система — единый неделимый «чёрный ящик» на L1", "Выполнено: устройство системы на уровне L1 не раскрывается"],
      ["Центральный блок и не менее 3 пользовательских ролей", "Выполнено: 3 персоны"],
      ["Не менее 3 внешних систем вне зоны контроля разработчиков", "Выполнено: 4 внешние системы"],
      ["Все связи строго направленные, с описанием данных", "Выполнено: 7 направленных связей, каждая с описанием и протоколом"],
      ["Декомпозиция центрального блока на контейнеры (L2)", "Выполнено: 5 контейнеров"],
      ["Не менее 5 контейнеров, включая мобильное и веб-приложение, бэкенд, хранилища", "Выполнено: все категории присутствуют; мобильное приложение — контейнер To-Be (обоснование в п. 5.3)"],
      ["Технологический стек для каждого контейнера", "Выполнено: приведён в таблице 3"],
      ["Сетевые протоколы для всех межсервисных связей", "Выполнено: указаны у каждой связи"],
      ["Таблица описания контейнеров", "Выполнено: п. 5.1"],
      ["Инструмент — PlantUML + C4-PlantUML", "Выполнено: макросы Person, System, System_Ext, Container, ContainerDb, Rel, LAYOUT_WITH_LEGEND, LAYOUT_LEFT_RIGHT"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: нарушений нет"],
    ],
    [55, 45],
  ),

  h1("8. Выводы"),
  body("По итогам лабораторной работы отработаны два первых уровня абстракции C4 Model на материале курсового проекта. Диаграмма системного контекста закрепила внешние границы системы DemoScope: три пользовательские роли, четыре внешние системы, семь направленных потоков данных с явными протоколами. Диаграмма контейнеров выполнила декомпозицию системы на пять независимо развёртываемых сущностей с фиксацией технологического стека каждого контейнера. Обе схемы описаны как код (Diagrams-as-Code) в нотации C4-PlantUML: текстовое описание диаграмм хранится в репозитории и автоматически превращается в изображения, что обеспечивает воспроизводимость и контроль версий архитектуры. Построенные уровни создают основу для следующего шага — декомпозиции Backend API на компоненты (Level 3) и последующего проектирования классов (Level 4)."),
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
