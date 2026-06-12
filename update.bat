@echo off
cd /d C:\sales_tracking

echo.
echo [1/3] Processing data...
set BAT_RUN=1
python run.py
if %errorlevel% neq 0 (
    echo ERROR: run.py failed.
    pause
    exit /b 1
)

echo.
echo [2/3] Staging files...
git add output/index.html
if %errorlevel% neq 0 (
    echo ERROR: git add failed.
    pause
    exit /b 1
)

for /f "tokens=1-3 delims=/" %%%%a in ("%date%") do (
    set TODAY=%%%%c-%%%%a-%%%%b
)
git commit -m "update %TODAY%"
if %errorlevel% neq 0 (
    echo No changes to commit.
    pause
    exit /b 0
)

echo.
echo [3/3] Pushing to GitHub...
git push
if %errorlevel% neq 0 (
    echo ERROR: git push failed.
    pause
    exit /b 1
)

echo.
echo Done! Dashboard will be updated in 1-2 minutes.
echo https://YJ-lee-0416.github.io/sales-dashboard/
echo.
pause
