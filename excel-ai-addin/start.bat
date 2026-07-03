@echo off
chcp 65001 >nul
title Excel AI 审计助手
echo =============================================
echo   Excel AI 审计助手 — 本地服务启动
echo =============================================
echo.
echo   [1/2] 生成图标...
python server.py --ensure-icons 2>nul
if %ERRORLEVEL% neq 0 (
    python -c "import server; server.ensure_icons()"
)
echo   图标就绪
echo.
echo   [2/2] 启动服务 (http://localhost:3000)...
echo.
echo   请在 Excel 中上传 manifest.xml 旁加载
echo   按 Ctrl+C 停止服务
echo.
echo =============================================
python server.py
pause
