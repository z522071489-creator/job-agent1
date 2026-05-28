@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo   Job Collector - Push to GitHub
echo ========================================
echo.

cd /d F:\WorkBuddy\job-collector

if not exist .git (
    echo [1/6] git init...
    git init
    if errorlevel 1 (
        echo ERROR: git init failed!
        pause
        exit /b 1
    )
    echo [1/6] git init OK
) else (
    echo [1/6] git already initialized
)

echo.
echo [2/6] git add ...
git add .
if errorlevel 1 (
    echo ERROR: git add failed!
    pause
    exit /b 1
)
echo [2/6] git add OK

echo.
echo [3/6] Checking git config...
git config user.name >nul 2>&1
if errorlevel 1 (
    echo Please enter your GitHub username:
    set /p uname=
    echo Please enter your GitHub email:
    set /p uemail=
    git config user.name "%uname%"
    git config user.email "%uemail%"
)
echo [3/6] config OK

echo.
echo [4/6] git commit...
git commit -m "feat: init job collector system"
if errorlevel 1 (
    echo WARNING: nothing to commit or commit failed, continuing...
)
echo [4/6] commit OK

echo.
echo [5/6] Set remote origin...
git remote remove origin >nul 2>&1
git remote add origin https://github.com/z522071489-creator/job-agent1.git
git branch -M main
echo [5/6] remote OK

echo.
echo [6/6] Pushing to GitHub...
echo This may ask for login. Please use browser to authorize if prompted.
echo.
git push -u origin main
if errorlevel 1 (
    echo.
    echo ========================================
    echo   PUSH FAILED!
    echo ========================================
    echo Possible reasons:
    echo   1. Wrong password - check your GitHub login
    echo   2. Network error - try again
    echo   3. Need to generate Personal Access Token
    echo      Go to GitHub Settings -> Developer Settings -> Tokens
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   ALL DONE! Code pushed successfully!
echo ========================================
pause
