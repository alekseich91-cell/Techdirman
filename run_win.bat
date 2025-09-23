@echo off
rem Автоматический запуск (Windows): создаёт виртуальное окружение, ставит зависимости и запускает приложение.

rem Получаем путь к каталогу текущего файла
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

rem Проверяем наличие Python
where python >nul 2>&1
if errorlevel 1 (
  echo Не найден python. Установите Python 3.10+ и убедитесь, что он доступен в PATH.
  goto :eof
)

rem Имя папки для виртуального окружения
set VENV_DIR=.venv

rem Создаём виртуальное окружение при отсутствии
if not exist "%VENV_DIR%" (
  python -m venv "%VENV_DIR%"
)

rem Активируем окружение
call "%VENV_DIR%\Scripts\activate.bat"

rem Обновляем pip и устанавливаем зависимости
python -m pip install --upgrade pip
pip install -r requirements.txt

rem Запускаем приложение
python TechDirRentMan\main.py