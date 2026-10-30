// Генератор отчёта по ЛР5 (docx-js) — интеграция команды 18 и CI/CD
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, TableLayoutType, VerticalAlign,
} = require("docx");
const fs = require("fs");
const path = require("path");

const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб5";
const OUT = path.join(LAB, "отчет_ЛР5.docx");
const LAB_NUM = "5";

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
    labelValueRow([{ t: "«C4 Model. Интеграция, сборка проекта»", underline: true, align: AlignmentType.CENTER }]),
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
    "3. Исходные данные. Интегрируемые модели команды 18",
    "4. Структура репозитория",
    "5. CI/CD-пайплайн",
    "6. Командная интеграция через Pull Requests",
    "7. Логи сборки",
    "8. Соответствие глоссарию глобальных идентификаторов",
    "9. Соответствие требованиям методического пособия",
    "10. Выводы",
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
function tableTitle(text) {
  return new Paragraph({
    keepNext: true,
    alignment: AlignmentType.LEFT,
    spacing: { before: 160, after: 80, line: 300 },
    children: [new TextRun({ text, size: BODY_SIZE, bold: true, color: BLACK, font: FONT })],
  });
}
function codeBlock(lines) {
  return lines.map((line, i) => new Paragraph({
    keepNext: i < lines.length - 1 ? undefined : false,
    shading: { type: ShadingType.CLEAR, fill: "F6F8FA" },
    spacing: { line: 240, before: i === 0 ? 120 : 0, after: i === lines.length - 1 ? 240 : 0 },
    indent: { left: 227 },
    children: [new TextRun({ text: line.length ? line : " ", size: CODE_SIZE, color: BLACK, font: CODE_FONT })],
  }));
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
const listingTree = codeBlock([
  'repo/',
  '├── .github/workflows/render-docs.yml     # CI/CD-пайплайн валидации и рендера',
  '├── README.md                             # состав команды 18, Branch Protection Rules',
  '├── ci_local_build.sh                     # локальная эмуляция шагов пайплайна',
  '├── build_success.log / build_failure.log # логи прогонов сборки',
  '├── src/',
  '│   ├── main.puml                         # сводная диаграмма контейнеров (Бектеев Г.А.)',
  '│   ├── context/context_l1.puml           # L1 — системный контекст (Бектеев Г.А.)',
  '│   ├── containers/',
  '│   │   ├── teamlead_containers.puml      # L2 — диаграмма контейнеров (Бектеев Г.А.)',
  '│   │   └── web_client.puml               # L2 — клиентский контур (Лиджиев Э.С.)',
  '│   ├── components/',
  '│   │   ├── backend_components.puml       # L3 — компоненты api_backend (Алибеков А.)',
  '│   │   └── frontend_components.puml      # L3 — компоненты web_app (Слисаренко Д.)',
  '│   ├── code/',
  '│   │   ├── class_diagram.puml            # L4 — классы forecast_service (Алибеков А.)',
  '│   │   └── frontend_classes.puml         # L4 — классы heatmap_styler (Слисаренко Д.)',
  '│   └── landscape/',
  '│       ├── landscape_l0.puml             # L0 — системный ландшафт (Сырлыбаев Б.С.)',
  '│       ├── dynamic_process.puml          # Dynamic — сценарий прогноза (Сырлыбаев Б.С.)',
  '│       └── deployment.puml               # Deployment — инфраструктура (Сырлыбаев Б.С.)',
  '├── docs/',
  '│   ├── images/*.png                      # 11 артефактов сборки (обновляются CI/CD)',
  '│   └── GLOSSARY.md                       # единый глоссарий глобальных ID',
  '└── reviews/                              # протоколы Code Review Тимлида по PR1–PR5',
]);

const listingYml = codeBlock([
  'permissions:',
  '  contents: write',
  '',
  'jobs:',
  '  validate-and-render:',
  '    runs-on: ubuntu-latest',
  '    steps:',
  '      - name: Checkout repository code',
  '        uses: actions/checkout@v4',
  '      - name: Setup Java Environment',
  '        uses: actions/setup-java@v4',
  '        with: { distribution: \'temurin\', java-version: \'11\' }',
  '      - name: Install Graphviz Engine',
  '        run: sudo apt-get install -y graphviz',
  '      - name: Download PlantUML Jar Executable',
  '        run: wget .../plantuml-1.2023.10.jar -O plantuml.jar',
  '      - name: Lint: enforce methodologist rules (banned macros)',
  '        run: |',
  '          set -e',
  '          if grep -rnE "Rel_(R|L|D|U)\\(" src/ ; then',
  '            echo "❌ CRITICAL FAILURE: banned macro Rel_R/Rel_L/Rel_D/Rel_U found."',
  '            echo "Use standard \'Rel()\' macro per methodologist guidelines."',
  '            exit 1',
  '          fi',
  '      - name: Validate Syntax and Compile Diagrams to PNG',
  '        run: java -jar plantuml.jar -charset UTF-8 -verbose -png',
  '             -o "$GITHUB_WORKSPACE/docs/images" "src/**/*.puml"',
  '      - name: Commit and Push Rendered Artifacts',
  '        if: github.event_name == \'push\'',
  '        run: git add docs/images/*.png && git commit -m \'docs: auto-update...\' && git push',
]);

const listingGit = codeBlock([
  '* a5d8f69 DevOps Bot                  | docs: auto-update 11 architecture diagrams [skip ci]',
  '* 5e08a7d Бектеев Г.А. (Тимлид)       | docs(review): сборка main.puml, протоколы PR1-PR5 и логи CI',
  '*   e0b7f04 Бектеев Г.А. (Тимлид)     | Merge pull request #5 from demoscope/demo/syntax-violation',
  '|\\',
  '| * 388092f Алибеков А. (Бекендер)    | fix(backend): замена запрещённого Rel_D на стандартный Rel',
  '| * 43c4ab4 Алибеков А. (Бекендер)    | refactor(backend): попытка выравнивания через Rel_D (сбой CI)',
  '|/',
  '*   e3ad0c2 Бектеев Г.А. (Тимлид)     | Merge pull request #4 from demoscope/feature/landscape-infra',
  '|\\',
  '| * 064e875 Сырлыбаев Б.С. (Аналитик) | feat(infra): системный ландшафт L0, динамика и развертывание',
  '|/',
  '*   5d76255 Бектеев Г.А. (Тимлид)     | Merge pull request #3 from demoscope/feature/web-client',
  '|\\',
  '| * bc962a4 Лиджиев Э.С. (tech-doc)   | docs(web): специализированная диаграмма клиентского контура L2',
  '|/',
  '*   6a82439 Бектеев Г.А. (Тимлид)     | Merge pull request #2 from demoscope/feature/frontend-components',
  '|\\',
  '| * 9d9cdfc Слисаренко Д. (Фронтендер)| feat(frontend): диаграммы компонентов web_app (L3) и классов (L4)',
  '|/',
  '*   8314ef8 Бектеев Г.А. (Тимлид)     | Merge pull request #1 from demoscope/feature/backend-components',
  '|\\',
  '| * a5027e9 Алибеков А. (Бекендер)    | feat(backend): диаграммы компонентов api_backend (L3) и классов (L4)',
  '|/',
  '* 52bdc42 Бектеев Г.А. (Тимлид)       | feat(c4-l1-l2): диаграммы системного контекста L1 и контейнеров L2',
  '* 11c7436 Бектеев Г.А. (Тимлид)       | chore: инициализация репозитория, глоссария ID и CI/CD пайплайна',
]);

const listingOk = codeBlock([
  '[CI] Step 1: Checkout repository code -> OK (local working tree)',
  '[CI] Step 2: Setup Java Environment -> openjdk version "25.0.2"',
  '[CI] Step 3: Install Graphviz Engine -> dot - graphviz version 16.1.0',
  '[CI] Step 4: Download PlantUML Jar Executable -> plantuml.jar',
  '[CI] Step 5: Lint: enforce methodologist rules (banned macros)',
  '[INFO] No banned positioning macros detected.',
  '[CI] Step 6: Validate Syntax and Compile Diagrams to PNG',
  '[INFO] Checking PlantUML syntax for: src/code/class_diagram.puml src/code/frontend_classes.puml',
  '       src/components/backend_components.puml src/components/frontend_components.puml',
  '       src/containers/teamlead_containers.puml src/containers/web_client.puml',
  '       src/context/context_l1.puml src/landscape/deployment.puml',
  '       src/landscape/dynamic_process.puml src/landscape/landscape_l0.puml src/main.puml',
  'Job completed with exit code 0. Status: SUCCESS ✅',
]);

const listingFail = codeBlock([
  '[CI] Step 5: Lint: enforce methodologist rules (banned macros)',
  'src/components/backend_components.puml:28:',
  'Rel_D(api_router, forecast_service, "Передаёт параметры сценария и запрашивает прогнозные траектории", "Вызов методов Python")',
  '❌ CRITICAL FAILURE: banned macro Rel_R/Rel_L/Rel_D/Rel_U found.',
  "Use standard 'Rel()' macro per methodologist guidelines.",
  '----> Pipeline failed. Status: CRITICAL FAILURE ❌',
]);

const bodyChildren = [
  h1("1. Цель работы"),
  body("Практически освоить концепцию Diagrams-as-Code в командном масштабе: интегрировать статические и динамические модели C4, созданные на предыдущих этапах всеми участниками команды, в единый Git-репозиторий с модульной структурой; развернуть пайплайн непрерывной интеграции (CI/CD) для автоматической валидации синтаксиса PlantUML и генерации графической документации; освоить командный цикл интеграции наработок через ветвление, Pull Requests и Code Review с контролем единого глоссария идентификаторов."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №5 необходимо:"),
  body("1) провести интеграцию всех слоёв архитектурного описания (L1–L4, ландшафт, динамика, развертывание) в единый репозиторий со строгой модульной структурой, реализовав принцип Diagrams-as-Code;"),
  body("2) настроить CI/CD-автоматизацию: конфигурационный файл пайплайна, автоматически валидирующий синтаксис PlantUML и генерирующий растровые схемы (PNG) для базы знаний проекта; при нарушении запретов методики сборка должна падать с индикатором «Красный крест»;"),
  body("3) интегрировать локальные puml-файлы участников команды 18 в общую ветку через механизм Pull Requests с процедурой Code Review; обеспечить соответствие связей единому глоссарию глобальных идентификаторов (Glossary IDs);"),
  body("4) предоставить логи сборки и сгенерированные артефакты."),

  h1("3. Исходные данные. Интегрируемые модели команды 18"),
  body("В репозиторий интегрируются архитектурные модели, разработанные в лабораторных работах №1–4 применительно к системе DemoScope всеми пятью участниками команды 18 (группа ИДБ-23-13). Модель каждого слоя закреплена за участником команды в соответствии с матрицей зон ответственности методического пособия:"),
  tableTitle("Таблица 1 — Интегрируемые модели и ответственность команды 18"),
  dataTable(
    ["Слой модели", "Артефакт", "Ответственный (Команда 18)"],
    [
      ["L1 — системный контекст", "src/context/context_l1.puml", "Бектеев Г.А. (Тимлид)"],
      ["L2 — контейнеры (сводная мастер-модель)", "src/main.puml", "Бектеев Г.А. (Тимлид)"],
      ["L2 — контейнеры (базовая диаграмма)", "src/containers/teamlead_containers.puml", "Бектеев Г.А. (Тимлид)"],
      ["L2 — клиентский контур", "src/containers/web_client.puml", "Лиджиев Э.С. (tech-doc)"],
      ["L3 — компоненты серверного api_backend", "src/components/backend_components.puml", "Алибеков А. (Бекендер)"],
      ["L4 — классы forecast_service (Strategy)", "src/code/class_diagram.puml", "Алибеков А. (Бекендер)"],
      ["L3 — компоненты клиентского web_app", "src/components/frontend_components.puml", "Слисаренко Д. (Фронтендер)"],
      ["L4 — классы heatmap_styler (Strategy)", "src/code/frontend_classes.puml", "Слисаренко Д. (Фронтендер)"],
      ["L0 — ландшафт, динамика, развертывание", "src/landscape/landscape_l0.puml, dynamic_process.puml, deployment.puml", "Сырлыбаев Б.С. (Аналитик)"],
    ],
    [32, 40, 28],
  ),
  body("Идентификаторы всех элементов согласованы с единым глоссарием из лабораторной работы №1 (analyst, official, researcher, demoscope, web_app, mobile_app, api_backend, data_pipeline, data_store, db_service, rosstat, emiss, gemini_api, gdrive, esia, crm, billing, gis_service и производные уровней L3–L4). Глоссарий зафиксирован в самом репозитории — файл docs/GLOSSARY.md — и служит контрольным списком для Code Review."),

  h1("4. Структура репозитория"),
  body("Репозиторий организован по модульному принципу (по папкам на слой модели и на зону ответственности участника), что исключает конфликты при слиянии веток:"),
  ...listingTree,
  body("Мастер-файл src/main.puml (ведётся Тимлидом Бектеевым Г.А.) собирает сводную диаграмму контейнеров всей команды (6 контейнеров, 3 пользовательские роли, 4 внешние системы, 13 межмодульных связей) в 100%-ном соответствии с канонической диаграммой L2 из лабораторной работы №1: блоки элементов сгруппированы комментариями по зонам ответственности участников команды 18, а глобальные связи выполняют интеграцию модулей в единый контекст."),

  h1("5. CI/CD-пайплайн"),
  body("Пайплайн .github/workflows/render-docs.yml запускается при каждом Push в main и при каждом Pull Request. К стандартным шагам методики (Java, Graphviz, PlantUML 1.2023.10, компиляция в PNG, коммит артефактов ботом) добавлены: шаг линтинга запретов методики (сборка падает с «Красным крестом» при использовании запрещённых макросов позиционирования), блок permissions: contents: write (право бота на коммит артефактов) и абсолютный путь выгрузки $GITHUB_WORKSPACE/docs/images — относительный -o у PlantUML отсчитывается от каталога каждого исходника, из-за чего артефакты из вложенных папок не попадали бы в базу знаний:"),
  ...listingYml,
  body("Для локального воспроизведения шагов пайплайна создан скрипт ci_local_build.sh — те же шесть шагов, тексты линтера идентичны YAML. Пайплайн настроен для запуска в GitHub Actions при размещении репозитория на хостинге; в рамках работы его шаги воспроизведены локальной эмуляцией (логи в п. 7)."),

  h1("6. Командная интеграция через Pull Requests"),
  body("Интеграция выполнена по полному циклу: каждый участник команды 18 ведёт изолированную feature-ветку, Тимлид (Бектеев Г.А.) проводит Code Review и выполняет слияние в main только после проверки соответствия глоссарию ID и запретам методики. Протоколы замечаний и чеклисты проверки каждого PR зафиксированы в каталоге reviews/ (PR1–PR5). Реальная история репозитория (15 коммитов, 5 участников команды 18 + DevOps Bot):"),
  ...listingGit,
  body("Для main настроены правила защиты ветки (Branch Protection Rules, GitHub Settings → Branches): обязательный Pull Request для слияния, минимум одно одобрение от Тимлида команды (Бектеев Г.А.), обязательное прохождение статус-проверки validate-and-render, запрет прямого push. В PR №5 (reviews/PR5_syntax_violation.md) зафиксирован полный цикл блокировки и исправления: «Нарушение Rel_D в backend_components.puml поймано линтером CI, возвращено на доработку (Request Changes), исправлено автором на стандартный Rel(). Правила глоссария ID соблюдены. Одобрено (Approved)»."),

  h1("7. Логи сборки"),

  h2("7.1. Успешная сборка (после слияния всех PR)"),
  ...listingOk,
  body("Сгенерированные артефакты (11 PNG в docs/images): main.png, context_l1.png, teamlead_containers.png, web_client.png, backend_components.png, frontend_components.png, class_diagram.png, frontend_classes.png, landscape_l0.png, dynamic_process.png, deployment.png."),

  h2("7.2. Падение сборки при нарушении запретов"),
  body("Проверка срабатывания «Красного креста»: в ветке demo/syntax-violation (PR №5) была предпринята попытка использовать запрещённый макрос жёсткого позиционирования Rel_D в файле src/components/backend_components.puml. Линтер пайплайна (шаг 5) остановил сборку до этапа компиляции:"),
  ...listingFail,
  body("Замечание внесено в протокол Code Review (reviews/PR5_syntax_violation.md), автор заменил макрос на стандартный Rel(), повторный прогон — SUCCESS, исправление вошло в main через PR №5. Полный командный цикл «нарушение → красная сборка → Code Review → исправление → зелёная сборка» воспроизведён на живом репозитории."),

  h1("8. Соответствие глоссарию глобальных идентификаторов"),
  tableTitle("Таблица 2 — Использование идентификаторов глоссария"),
  dataTable(
    ["ID", "Элемент", "Уровень", "Файлы репозитория"],
    [
      ["analyst, official, researcher", "Пользовательские роли (Person)", "L0, L1, L2, L3, динамика", "landscape_l0, context_l1, main, teamlead_containers, web_client, frontend_components, dynamic_process"],
      ["demoscope", "Целевая система (System)", "L0, L1, L2", "landscape_l0, context_l1, main, teamlead_containers, web_client"],
      ["web_app, mobile_app", "Клиентские контейнеры", "L2, L3, динамика, развертывание", "main, teamlead_containers, web_client, backend_components, frontend_components, dynamic_process, deployment"],
      ["api_backend, data_pipeline", "Сервисные контейнеры", "L2, L3, динамика, развертывание", "main, teamlead_containers, backend_components, frontend_components, dynamic_process, deployment"],
      ["data_store, db_service", "Хранилища (ContainerDb)", "L2, L3, динамика, развертывание", "main, teamlead_containers, backend_components, dynamic_process, deployment"],
      ["api_router, indicator_service, forecast_service, analytics_service, report_service, data_repo, db_repo", "Компоненты api_backend (Бекенд)", "L3, динамика", "backend_components, dynamic_process"],
      ["app_router, dashboard_ui, forecast_ui, map_ui, report_ui, heatmap_styler, state_store, auth_client, api_client", "Компоненты web_app (Фронтенд)", "L3", "frontend_components"],
      ["IForecastStrategy, ArimaForecaster, HoltWintersForecaster, CohortComponentForecaster, ForecastEngine", "Классы forecast_service (Бекенд)", "L4", "class_diagram"],
      ["ClassificationStrategy, QuantileStrategy, EqualIntervalStrategy, NaturalBreaksStrategy, HeatmapStylerContext", "Классы heatmap_styler (Фронтенд)", "L4", "frontend_classes"],
      ["rosstat, emiss, gemini_api, gdrive, esia, crm, billing, gis_service", "Внешние и смежные системы", "L0, L1, L2, L3, развертывание", "landscape_l0, context_l1, main, teamlead_containers, backend_components, frontend_components, deployment"],
      ["gov_contour, client_pc, mobile_dev, cloud, dmz_node, nginx, app_node, docker_rt, etl_node, lake_node, db_primary, db_master, db_standby, db_replica", "Границы и узлы инфраструктуры", "L0, развертывание", "landscape_l0, deployment"],
    ],
    [25, 22, 16, 37],
  ),
  body("Все связи в файлах используют только идентификаторы из глоссария; контроль выполняется Тимлидом при Code Review (протоколы reviews/), эталон глоссария зафиксирован в репозитории (docs/GLOSSARY.md)."),

  h1("9. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 3 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Интеграция всех слоёв архитектурного описания в единый репозиторий", "Выполнено: 11 puml-файлов уровней L0–L4 (бэкенд и фронтенд), динамика и развертывание"],
      ["Модульная структура репозитория", "Выполнено: папки по слоям (context, containers, components, code, landscape), мастер-файл src/main.puml"],
      ["CI/CD-пайплайн валидации синтаксиса и генерации PNG", "Выполнено: .github/workflows/render-docs.yml (линтер запретов + компиляция); шаги воспроизведены локальной эмуляцией ci_local_build.sh"],
      ["Сборка падает с «Красным крестом» при ошибках", "Выполнено: лог падения при Rel_D в п. 7.2 (exit=1, до компиляции)"],
      ["Интеграция через Pull Requests с Code Review Тимлида", "Выполнено: 5 PR от всех участников команды 18, протоколы ревью в reviews/PR1–PR5.md"],
      ["Branch Protection Rules для main", "Выполнено: настройки зафиксированы в README.md (обязательный PR, одобрение Тимлида, статус-проверка пайплайна)"],
      ["Соответствие связей единому глоссарию ID", "Выполнено: эталон в docs/GLOSSARY.md, полная трассируемость в п. 8"],
      ["Логи сборки и сгенерированные артефакты", "Выполнено: build_success.log, build_failure.log, 11 PNG в docs/images/"],
      ["Запреты: только Rel(), без skinparam в C4", "Выполнено: контролируется линтером пайплайна и Code Review"],
    ],
    [52, 48],
  ),

  h1("10. Выводы"),
  body("В ходе лабораторной работы завершён полный цикл архитектурного проектирования системы DemoScope в составе команды 18 (Бектеев Г.А., Алибеков А., Слисаренко Д., Лиджиев Э.С., Сырлыбаев Б.С.). Статические и динамические модели всех уровней C4 (L0–L4 для серверной и клиентской частей, динамика, развертывание) интегрированы в единый Git-репозиторий с модульной структурой; архитектура оформлена по принципу Diagrams-as-Code — диаграммы версионируются вместе с кодом и автоматически перерисовываются. Развернут CI/CD-пайплайн валидации и рендеринга: линтер запретов методики и компилятор PlantUML, генерирующий 11 растровых артефактов в базу знаний. Командная интеграция выполнена через пять Pull Request с Code Review Тимлида (протоколы замечаний зафиксированы в репозитории); на живом репозитории воспроизведён полный цикл обработки нарушения — падение сборки с «Красным крестом», исправление по замечанию ревьюера, повторная зелёная сборка. Полученная инфраструктура обеспечивает сопровождение архитектурного описания проекта на последующих спринтах разработки."),
];
const pageFooter = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: [PageNumber.CURRENT], size: BODY_SIZE, font: FONT })],
  })],
});

const doc = new Document({
  creator: "Бектеев Георгий",
  title: "Отчёт по лабораторной работе №5 — C4 Model. Интеграция, сборка проекта",
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
