@echo off
chcp 65001 >nul
echo ========================================
echo   招聘采集系统 - 一键推送到 GitHub
echo ========================================
echo.
echo [1/6] 进入项目目录...
cd /d F:\WorkBuddy\job-collector
echo.
echo [2/6] 初始化 Git...
git init
echo.
echo [3/6] 添加文件...
git add .
echo.
echo [4/6] 提交代码...
git commit -m "feat: 招聘采集系统初始化"
echo.
echo [5/6] 关联远程仓库...
git remote add origin https://github.com/z522071489-creator/job-agent1.git 2>nul
git branch -M main
echo.
echo [6/6] 推送到 GitHub...
git push -u origin main
echo.
echo ========================================
echo   完成！按任意键关闭窗口
echo ========================================
pause >nul
