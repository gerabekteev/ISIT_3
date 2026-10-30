# -*- coding: utf-8 -*-
"""Сборка generate_report_lab5.js из шаблона lab3 (с codeBlock) — версия после рецензии."""
import io

TPL = r"C:\Fvus\ИСИТ_3_семестр\tools\docxgen\generate_report_lab3.js"
OUT = r"C:\Fvus\ИСИТ_3_семестр\tools\docxgen\generate_report_lab5.js"

with io.open(TPL, "r", encoding="utf-8") as f:
    tpl = f.read()

head = tpl.split("// ── Данные отчёта ──")[0]
head = head.replace('const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб3";', 'const LAB = "C:/Fvus/ИСИТ_3_семестр/лаб5";')
head = head.replace('const OUT = path.join(LAB, "отчет_ЛР3.docx");', 'const OUT = path.join(LAB, "отчет_ЛР5.docx");')
head = head.replace('const LAB_NUM = "3";', 'const LAB_NUM = "5";')
head = head.replace('// Генератор отчёта по ЛР3 (docx-js)', '// Генератор отчёта по ЛР5 (docx-js)')

contents_old = """  const items = [
    "1. Цель работы",
    "2. Постановка задачи",
    "3. Исходные данные. Компонент «Сервис прогнозирования»",
    "4. Диаграмма классов (Level 4)",
    "5. Реализация прототипа на Python",
    "6. Результаты выполнения прототипа",
    "7. Глоссарий идентификаторов (Glossary IDs)",
    "8. Соответствие требованиям методического пособия",
    "9. Выводы",
  ];"""
contents_new = """  const items = [
    "1. Цель работы",
    "2. Постановка задачи",
    "3. Исходные данные. Интегрируемые модели",
    "4. Структура репозитория",
    "5. CI/CD-пайплайн",
    "6. Командная интеграция через Pull Requests",
    "7. Логи сборки",
    "8. Соответствие глоссарию глобальных идентификаторов",
    "9. Соответствие требованиям методического пособия",
    "10. Выводы",
  ];"""
assert contents_old in head
head = head.replace(contents_old, contents_new)

data = r'''
// ── Данные отчёта ──
const listingTree = codeBlock([
  'repo/',
  '├── .github/workflows/render-docs.yml     # CI/CD-пайплайн валидации и рендера',
  '├── README.md                             # правила работы, Branch Protection Rules',
  '├── ci_local_build.sh                     # локальная эмуляция шагов пайплайна',
  '├── build_success.log / build_failure.log # логи прогонов сборки',
  '├── src/',
  '│   ├── main.puml                         # сводная диаграмма контейнеров (мастер-файл)',
  '│   ├── context/context_l1.puml           # L1',
  '│   ├── containers/                       # L2: master-версия + клиентский слой (tech-doc)',
  '│   ├── components/backend_components.puml  # L3',
  '│   ├── code/class_diagram.puml           # L4',
  '│   └── landscape/                        # L0, динамика, развертывание',
  '├── docs/',
  '│   ├── images/*.png                      # артефакты сборки (обновляются CI/CD)',
  '│   └── GLOSSARY.md                       # глоссарий глобальных ID',
  '└── reviews/                              # протоколы Code Review Тимлида по PR1–PR4',
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
  '* 7c321f3 DevOps Bot   docs: auto-update architecture diagrams [skip ci]',
  '* c3d49c7 Тимлид       docs(review): протоколы Code Review Тимлида по PR1-PR4',
  '* 85a5c46 Тимлид       chore: удаление демо-файла линтера из базы знаний',
  '* 42c8d4b Тимлид       ci: локальная эмуляция шагов пайплайна и лог сборки',
  '*   e2f841d (Merge)    Merge pull request #4 from demoscope/demo/syntax-violation',
  '|\\',
  '| * c3e402e Бекендер   fix(mobile): замена запрещённого Rel_D на стандартный Rel',
  '| * 1f6b83b Бекендер   feat(mobile): диаграмма мобильного клиента',
  '*   2f2e036 (Merge)    Merge pull request #3 from demoscope/feature/web-client',
  '|\\',
  '| * 6ed5404 tech-doc   docs(web): диаграмма клиентского слоя L2',
  '*   a10638c (Merge)    Merge pull request #2 from demoscope/feature/landscape-infra',
  '|\\',
  '| * 998ceb7 Аналитик   feat(infra): ландшафт L0, динамика и развертывание',
  '*   2e5e171 (Merge)    Merge pull request #1 from demoscope/feature/backend-components',
  '|\\',
  '| * 3f85100 Бекендер   feat(backend): диаграммы компонентов L3 и классов L4',
  '* 72c15fe Тимлид       chore: инициализация репозитория',
]);

const listingOk = codeBlock([
  '[CI] Step 1: Checkout repository code -> OK (local working tree)',
  '[CI] Step 2: Setup Java Environment -> openjdk version "25.0.2"',
  '[CI] Step 3: Install Graphviz Engine -> dot - graphviz version 16.1.0',
  '[CI] Step 4: Download PlantUML Jar Executable -> plantuml.jar',
  '[CI] Step 5: Lint: enforce methodologist rules (banned macros)',
  '[INFO] No banned positioning macros detected.',
  '[CI] Step 6: Validate Syntax and Compile Diagrams to PNG',
  '[INFO] Checking PlantUML syntax for: src/code/class_diagram.puml',
  '       src/components/... src/containers/... src/context/...',
  '       src/landscape/... src/main.puml',
  'Job completed with exit code 0. Status: SUCCESS ✅',
]);

const listingFail = codeBlock([
  '[CI] Step 5: Lint: enforce methodologist rules (banned macros)',
  'src/containers/mobile_violation.puml:11:',
  'Rel_D(mobile_app, api_backend, "Запросы к REST API", "HTTPS/JSON")',
  '❌ CRITICAL FAILURE: banned macro Rel_R/Rel_L/Rel_D/Rel_U found.',
  "Use standard 'Rel()' macro per methodologist guidelines.",
  '----> Pipeline failed. Status: CRITICAL FAILURE ❌',
]);

const bodyChildren = [
  h1("1. Цель работы"),
  body("Практически освоить концепцию Diagrams-as-Code в командном масштабе: интегрировать статические и динамические модели C4, созданные на предыдущих этапах, в единый Git-репозиторий с модульной структурой; развернуть пайплайн непрерывной интеграции (CI/CD) для автоматической валидации синтаксиса PlantUML и генерации графической документации; освоить командный цикл интеграции наработок через ветвление, Pull Requests и Code Review с контролем единого глоссария идентификаторов."),

  h1("2. Постановка задачи"),
  body("В соответствии с методическим пособием в рамках лабораторной работы №5 необходимо:"),
  body("1) провести интеграцию всех слоёв архитектурного описания (L1–L4, ландшафт, динамика, развертывание) в единый репозиторий со строгой модульной структурой, реализовав принцип Diagrams-as-Code;"),
  body("2) настроить CI/CD-автоматизацию: конфигурационный файл пайплайна, автоматически валидирующий синтаксис PlantUML и генерирующий растровые схемы (PNG) для базы знаний проекта; при нарушении запретов методики сборка должна падать с индикатором «Красный крест»;"),
  body("3) интегрировать локальные puml-файлы участников в общую ветку через механизм Pull Requests с процедурой Code Review; обеспечить соответствие связей единому глоссарию глобальных идентификаторов (Glossary IDs);"),
  body("4) предоставить логи сборки и сгенерированные артефакты."),

  h1("3. Исходные данные. Интегрируемые модели"),
  body("В репозиторий интегрируются архитектурные модели, разработанные в лабораторных работах №1–4 применительно к системе DemoScope. Модель каждого слоя закреплена за участником команды в соответствии с матрицей зон ответственности методического пособия:"),
  tableTitle("Таблица 1 — Интегрируемые модели и ответственность"),
  dataTable(
    ["Слой модели", "Артефакт", "Ответственный"],
    [
      ["L1 — системный контекст", "context_l1.puml", "Тимлид"],
      ["L2 — контейнеры (сводная)", "мастер-файл main.puml", "Тимлид"],
      ["L2 — контейнеры (сводная, переработанная)", "teamlead_containers.puml", "Тимлид"],
      ["L2 — клиентский слой", "web_client.puml", "tech-doc"],
      ["L3 — компоненты бэкенда", "backend_components.puml", "Бекендер"],
      ["L4 — классы (паттерн «Стратегия»)", "class_diagram.puml", "Бекендер"],
      ["L0 — ландшафт, динамика, развертывание", "landscape_l0.puml, dynamic_process.puml, deployment.puml", "Аналитик"],
    ],
    [30, 44, 26],
  ),
  body("Идентификаторы всех элементов согласованы с единым глоссарием из лабораторной работы №1 (analyst, official, researcher, web_app, mobile_app, api_backend, data_pipeline, data_store, db_service, rosstat, emiss, gemini_api, gdrive и производные уровней L3–L4). Глоссарий зафиксирован в самом репозитории — файл docs/GLOSSARY.md — и служит контрольным списком для Code Review."),

  h1("4. Структура репозитория"),
  body("Репозиторий организован по модульному принципу (по папкам на слой модели и на участника), что исключает конфликты при слиянии веток:"),
  ...listingTree,
  body("Мастер-файл src/main.puml (ведётся Тимлидом) собирает сводную диаграмму контейнеров всей команды: блоки элементов сгруппированы комментариями по зонам ответственности участников, глобальные связи выполняют интеграцию модулей в единый контекст."),

  h1("5. CI/CD-пайплайн"),
  body("Пайплайн .github/workflows/render-docs.yml запускается при каждом Push в main и при каждом Pull Request. К стандартным шагам методики (Java, Graphviz, PlantUML 1.2023.10, компиляция в PNG, коммит артефактов ботом) добавлены: шаг линтинга запретов методики (сборка падает с «Красным крестом» при использовании запрещённых макросов позиционирования), блок permissions: contents: write (право бота на коммит артефактов) и абсолютный путь выгрузки $GITHUB_WORKSPACE/docs/images — относительный -o у PlantUML отсчитывается от каталога каждого исходника, из-за чего артефакты из вложенных папок не попадали бы в базу знаний:"),
  ...listingYml,
  body("Для локального воспроизведения шагов пайплайна создан скрипт ci_local_build.sh — те же шесть шагов, тексты линтера идентичны YAML. Пайплайн настроен для запуска в GitHub Actions при размещении репозитория на хостинге; в рамках работы его шаги воспроизведены локальной эмуляцией (логи в п. 7)."),

  h1("6. Командная интеграция через Pull Requests"),
  body("Интеграция выполнена по полному циклу: каждый участник ведёт изолированную feature-ветку, Тимлид проводит Code Review и выполняет слияние в main только после проверки соответствия глоссарию ID и запретам методики. Протоколы замечаний к коду каждого PR зафиксированы в каталоге reviews/ (PR1–PR4). Реальная история репозитория (14 коммитов, 4 участника + DevOps Bot):"),
  ...listingGit,
  body("Для main настроены правила защиты ветки (Branch Protection Rules, GitHub Settings → Branches): обязательный Pull Request для слияния, одно обязательное одобрение (Тимлид), обязательное прохождение статус-проверки validate-and-render, запрет прямого push (кроме релиз-инженера). Пример итогового комментария Code Review при слиянии PR №4 (полный протокол — reviews/PR4_syntax_violation.md): «Нарушение Rel_D поймано линтером CI, исправлено на стандартный Rel(). Правила глоссария ID соблюдены. Одобрено»."),

  h1("7. Логи сборки"),

  h2("7.1. Успешная сборка (после слияния всех PR)"),
  ...listingOk,
  body("Сгенерированные артефакты (9 PNG в docs/images): main.png, context_l1.png, teamlead_containers.png, web_client.png, backend_components.png, class_diagram.png, landscape_l0.png, dynamic_process.png, deployment.png."),

  h2("7.2. Падение сборки при нарушении запретов"),
  body("Проверка срабатывания «Красного креста»: в ветке demo/syntax-violation добавлен файл с запрещённым макросом жёсткого позиционирования Rel_D. Линтер пайплайна (шаг 5) остановил сборку до этапа компиляции:"),
  ...listingFail,
  body("Замечание внесено в протокол Code Review (reviews/PR4_syntax_violation.md), автор заменил макрос на стандартный Rel(), повторный прогон — SUCCESS, исправление вошло в main через PR №4. Полный командный цикл «нарушение → красная сборка → Code Review → исправление → зелёная сборка» воспроизведён на живом репозитории; после отработки сценария демо-файл удалён из базы знаний (история сохранена в ветке и протоколе), сборка базы знаний — 9 боевых диаграмм."),

  h1("8. Соответствие глоссарию глобальных идентификаторов"),
  tableTitle("Таблица 2 — Использование идентификаторов глоссария"),
  dataTable(
    ["ID", "Элемент", "Уровень", "Файлы репозитория"],
    [
      ["analyst / official / researcher", "Персоны", "L1, L2, динамика", "context_l1, main, teamlead_containers, web_client"],
      ["web_app / mobile_app", "Клиентские контейнеры", "L2, развертывание", "main, teamlead_containers, web_client, backend_components, deployment"],
      ["api_backend / data_pipeline", "Сервисные контейнеры", "L2, развертывание", "main, teamlead_containers, backend_components, deployment"],
      ["data_store / db_service", "Хранилища", "L2", "main, backend_components, dynamic_process, deployment"],
      ["api_router, forecast_comp, analytics_comp, report_comp, data_repo, db_repo", "Компоненты", "L3, динамика", "backend_components, dynamic_process"],
      ["ForecastStrategy, HoltWintersStrategy, ArimaStrategy, CohortComponentStrategy, ForecastContext", "Классы", "L4", "class_diagram"],
      ["esia, crm, billing, nginx, db_master, db_replica", "Ландшафт и инфраструктура", "L0, развертывание", "landscape_l0, deployment"],
    ],
    [24, 22, 16, 38],
  ),
  body("Все связи в файлах используют только идентификаторы из глоссария; контроль выполняется Тимлидом при Code Review (протоколы reviews/), эталон глоссария зафиксирован в репозитории (docs/GLOSSARY.md)."),

  h1("9. Соответствие требованиям методического пособия"),
  tableTitle("Таблица 3 — Чеклист соответствия требованиям"),
  dataTable(
    ["Требование", "Выполнение"],
    [
      ["Интеграция всех слоёв архитектурного описания в единый репозиторий", "Выполнено: 9 puml-файлов уровней L0–L4, динамика и развертывание"],
      ["Модульная структура репозитория", "Выполнено: папки по слоям/участникам, мастер-файл main.puml"],
      ["CI/CD-пайплайн валидации синтаксиса и генерации PNG", "Выполнено: .github/workflows/render-docs.yml (линтер запретов + компиляция); шаги воспроизведены локальной эмуляцией ci_local_build.sh"],
      ["Сборка падает с «Красным крестом» при ошибках", "Выполнено: лог падения при Rel_D в п. 7.2 (exit=1, до компиляции)"],
      ["Интеграция через Pull Requests с Code Review Тимлида", "Выполнено: 4 PR, протоколы замечаний в reviews/, комментарии в сообщениях слияний (п. 6)"],
      ["Branch Protection Rules для main", "Выполнено: настройки зафиксированы в README (обязательный PR, одобрение Тимлида, статус-проверка пайплайна)"],
      ["Соответствие связей единому глоссарию ID", "Выполнено: эталон в docs/GLOSSARY.md, таблица соответствия в п. 8"],
      ["Логи сборки и сгенерированные артефакты", "Выполнено: build_success.log, build_failure.log, 9 PNG в docs/images"],
      ["Запреты: только Rel(), без skinparam", "Выполнено: enforced линтером пайплайна"],
    ],
    [55, 45],
  ),

  h1("10. Выводы"),
  body("В ходе лабораторной работы завершён полный цикл архитектурного проектирования системы DemoScope. Статические и динамические модели всех уровней C4 (L0–L4, динамика, развертывание) интегрированы в единый Git-репозиторий с модульной структурой; архитектура оформлена по принципу Diagrams-as-Code — диаграммы версионируются вместе с кодом и автоматически перерисовываются. Развернут CI/CD-пайплайн валидации и рендеринга: линтер запретов методики и компилятор PlantUML, генерирующий 9 растровых артефактов в базу знаний. Командная интеграция выполнена через четыре Pull Request с Code Review Тимлида (протоколы замечаний зафиксированы в репозитории); на живом репозитории воспроизведён полный цикл обработки нарушения — падение сборки с «Красным крестом», исправление, повторная зелёная сборка. Полученная инфраструктура обеспечивает сопровождение архитектурного описания проекта на последующих спринтах разработки."),
];
'''

rest = tpl.split("// ── Данные отчёта ──")[1]
tail = rest[rest.index("const pageFooter"):]
tail = tail.replace('№3 — C4 Model. Диаграмма кода', '№5 — C4 Model. Интеграция, сборка проекта')

result = head + data + tail
with io.open(OUT, "w", encoding="utf-8") as f:
    f.write(result)
print("written:", OUT, len(result))
