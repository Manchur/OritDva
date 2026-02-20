@echo off
echo ============================================
echo   Building OritDva Windows App
echo ============================================
echo.

:: Check if PyInstaller is installed
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building executable...
echo.

pyinstaller ^
    --name "OritDva" ^
    --onedir ^
    --windowed ^
    --noconfirm ^
    --clean ^
    --add-data ".env.example;." ^
    --hidden-import "win32com" ^
    --hidden-import "win32com.client" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pywintypes" ^
    --hidden-import "win32timezone" ^
    --collect-all "google.genai" ^
    app_gui.py

echo.
if %errorlevel% equ 0 (
    echo ============================================
    echo   BUILD SUCCESSFUL!
    echo ============================================
    echo.
    echo Output: dist\OritDva\OritDva.exe
    echo.
    echo To distribute:
    echo   1. Copy the entire dist\OritDva folder
    echo   2. Create a samples\ folder inside it
    echo   3. Run OritDva.exe
    echo.

    :: Copy .env.example into dist
    if not exist "dist\OritDva\.env" (
        copy ".env.example" "dist\OritDva\.env" >nul 2>&1
    )

    :: Create samples dir in dist
    if not exist "dist\OritDva\samples" (
        mkdir "dist\OritDva\samples"
    )
) else (
    echo ============================================
    echo   BUILD FAILED - Check errors above
    echo ============================================
)

pause
