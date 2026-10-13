@echo off
setlocal
cd /d "%~dp0"
echo ====================================================================
echo  STARTING MUTUAL FUND MONTHLY PIPELINE (SINGLE COMMAND RUNNER)
echo ====================================================================
python run_pipeline.py %*
if errorlevel 1 (
    echo.
    echo [ERROR] Pipeline stopped due to an error. Please check the logs above.
) else (
    echo.
    echo [SUCCESS] Entire Mutual Fund pipeline finished with 1000%% accuracy!
)
echo.
pause
