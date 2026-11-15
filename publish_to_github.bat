@echo off
setlocal EnableExtensions
set "REPO_URL=https://github.com/realahmad07/NASA.git"
set "WORK_DIR=%~dp0..\NASA-GitHub-Publish"

where git >nul 2>nul
if errorlevel 1 (echo Install Git for Windows first.& exit /b 1)

if not exist "%WORK_DIR%\.git" (
  if exist "%WORK_DIR%" rmdir /s /q "%WORK_DIR%"
  git clone "%REPO_URL%" "%WORK_DIR%"
  if errorlevel 1 exit /b 1
) else (
  git -C "%WORK_DIR%" pull --ff-only origin main
  if errorlevel 1 exit /b 1
)

robocopy "%~dp0" "%WORK_DIR%" /E /XD .git .venv __pycache__ .pytest_cache /XF .env *.sqlite3 *.db *.log
if errorlevel 8 exit /b 1
cd /d "%WORK_DIR%"

git add server.py requirements.txt pyproject.toml .env.example .gitignore start_windows.bat start_unix.sh data/.gitkeep
git diff --cached --quiet || git commit -m "feat: publish PyroSage backend and runtime configuration"
if errorlevel 1 exit /b 1

git add static/index.html techniques.json reference_prototype.html
git diff --cached --quiet || git commit -m "feat: publish evidence dashboard and scientific technique registry"
if errorlevel 1 exit /b 1

git add README.md docs tests
git diff --cached --quiet || git commit -m "docs: publish technical documentation and verification suite"
if errorlevel 1 exit /b 1

git push origin main
if errorlevel 1 exit /b 1
echo PyroSage source, dashboard, tests, and documentation pushed successfully.
endlocal
