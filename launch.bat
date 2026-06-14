@echo off
chcp 65001 > nul
cls

echo ================================================
echo    ISO27001/ISO20000 检查表生成工具
echo ================================================
echo.

REM 获取脚本所在目录
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM 获取Python路径
set PYTHON_EXE=C:\Users\Confu\.workbuddy\binaries\python\versions\3.13.12\python.exe

echo 请填写以下参数：
echo.

set /p RECORDS_DIR="[1/4] 客户记录目录路径: "
set /p STANDARD="[2/4] 标准类型 (ISO27001 或 ISO20000): "
set /p AUDIT_DATE="[3/4] 审核日期 (格式: YYYY-MM-DD): "
set /p COMPANY_NAME="[4/4] 企业名称: "

echo.
echo ================================================
echo 开始生成检查表...
echo 详细日志将保存到 output 子目录
echo ================================================
echo.

REM 执行安静模式启动器
"%PYTHON_EXE%" -u "%SCRIPT_DIR%run_quiet.py" ^
    --records-dir "%RECORDS_DIR%" ^
    --standard %STANDARD% ^
    --audit-date %AUDIT_DATE% ^
    --company-name "%COMPANY_NAME%"

echo.
echo ================================================
echo 生成完成！
echo 输出文件在: %SCRIPT_DIR%output\
echo ================================================
echo.
pause
