@echo off
chcp 65001 >nul
title zhihu
cd /d "E:\Projects\my-workspace\zhihu-assistant"
cls

echo ==============================
echo      知乎助手
echo ==============================
echo.
echo   1 - 获取 Cookie
echo       （登录浏览器，扫码即可）
echo.
echo   2 - 运行全流程
echo       （自动抓热点 + 生成5篇草稿）
echo.
echo   3 - 查看草稿
echo       （浏览器打开草稿页）
echo.
echo   4 - 退出
echo.
echo ==============================
set /p ch="请输入选项 (1-4): "

if "%ch%"=="1" (cls && python get_cookie.py && echo. && pause && exit)
if "%ch%"=="2" (cls && python main.py && echo. && pause && exit)
if "%ch%"=="3" (cls && python view_drafts.py && start "" "drafts.html" && echo. && pause && exit)
if "%ch%"=="4" (exit)
echo 输入无效
pause
exit