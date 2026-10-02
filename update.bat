@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

echo [1/3] Building dashboard (input -^> output\index.html)...
python run.py
if errorlevel 1 (
  echo [ERROR] Build failed. Nothing was uploaded.
  pause
  exit /b 1
)

echo [2/3] Committing...
git add output/index.html run.py template_v2.html update.bat requirements.txt
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "update"
  if errorlevel 1 (
    echo [ERROR] Commit failed. Check git user.name / user.email settings.
    pause
    exit /b 1
  )
) else (
  echo No changes to upload.
  pause
  exit /b 0
)

echo [3/3] Uploading to GitHub...
git push origin main
if errorlevel 1 (
  echo [ERROR] Push failed. Check login or network.
  pause
  exit /b 1
)
echo Done. Dashboard refreshes in 1-2 minutes:
echo https://YJ-lee-0416.github.io/sales-dashboard/
pause