@echo off
echo ============================================
echo Compiling Student Performance Analytics...
echo ============================================
javac StudentPerformanceAnalytics.java
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Compilation failed. Please check your Java installation and code.
    pause
    exit /b %errorlevel%
)
echo Compilation successful.
echo.
echo ============================================
echo Running Student Performance Analytics Tool...
echo ============================================
java StudentPerformanceAnalytics
pause
