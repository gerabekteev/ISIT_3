@echo off
chcp 65001 >nul
setlocal
rem ============================================================
rem  Стартовый скрипт лабораторных работ по C4 Model
rem  1) Поднимает Docker (если выключен)
rem  2) Запускает локальный PlantUML Server (порт 8080)
rem  3) Рендерит все .puml из лаб1\diagrams в PNG (jar + Graphviz)
rem  4) Открывает Antigravity IDE (запасной вариант: VS Code)
rem ============================================================
set "ROOT=%~dp0"
set "PATH=%PATH%;C:\Program Files\Graphviz\bin"

echo [1/4] Проверка Docker...
docker info >nul 2>&1
if not errorlevel 1 goto docker_ok
echo     Docker Desktop не запущен - запускаю и жду (до 2 минут)...
start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"
set /a TRIES=0
:waitdocker
timeout /t 5 /nobreak >nul
docker info >nul 2>&1
if not errorlevel 1 goto docker_ok
set /a TRIES+=1
if %TRIES% lss 24 goto waitdocker
echo [ОШИБКА] Docker не поднялся за 2 минуты. Открой Docker Desktop вручную и запусти скрипт снова.
pause
exit /b 1
:docker_ok
echo     OK: Docker работает.

echo [2/4] Запуск локального PlantUML Server (порт 8090)...
docker compose -f "%ROOT%docker-compose.yml" up -d
set /a TRIES=0
:waitserver
curl -s -o nul http://localhost:8090
if not errorlevel 1 goto server_ok
set /a TRIES+=1
if %TRIES% lss 12 goto waitserver
echo [ПРЕДУПРЕЖДЕНИЕ] Сервер пока не отвечает - предпросмотр Alt+D заработает, когда контейнер прогрузится.
goto server_done
:server_ok
echo     OK: сервер доступен на http://localhost:8090
:server_done

echo [3/4] Рендеринг диаграмм (.puml -^> PNG)...
if not exist "%ROOT%лаб1\diagrams\*.puml" (
    echo     В лаб1\diagrams нет .puml файлов - рендер пропущен.
    goto render_done
)
java -jar "%ROOT%tools\plantuml.jar" -charset UTF-8 -failfast2 -png -o "%ROOT%лаб1\diagrams\png" "%ROOT%лаб1\diagrams\*.puml"
if errorlevel 1 (
    echo [ОШИБКА] Синтаксис одного из .puml неверен - см. сообщения выше ^(сборка с такой ошибкой в CI/CD упадет^).
) else (
    echo     OK: PNG лежат в лаб1\diagrams\png
)
:render_done

echo [4/4] Запуск IDE...
set "AG=%LOCALAPPDATA%\Programs\Antigravity IDE\bin\antigravity-ide.cmd"
if exist "%AG%" (
    start "" "%AG%" "%ROOT%"
    echo     OK: Antigravity IDE открыт.
) else (
    start "" code "%ROOT%"
    echo     OK: VS Code открыт ^(Antigravity не найден^).
)

echo.
echo ====== Всё готово. ======
echo Предпросмотр диаграммы в IDE: Alt+D  ^(рендер через локальный PlantUML Server^).
pause
