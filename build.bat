@echo off
chcp 65001 >nul
echo ========================================
echo 邮件导出工具自动打包脚本
echo ========================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python
    pause
    exit /b 1
)

:: 检查PyInstaller是否安装
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo 错误: PyInstaller安装失败
        pause
        exit /b 1
    )
)

:: 清理之前的构建文件
echo 清理之前的构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"

:: 使用PyInstaller打包GUI程序
echo.
echo 正在打包GUI程序...
pyinstaller --onefile --windowed --name="mail_exporter_gui" --add-data="README.md;." --add-data="LICENSE;." mail_exporter_gui.py
if errorlevel 1 (
    echo 错误: GUI程序打包失败
    pause
    exit /b 1
)

:: 使用PyInstaller打包命令行程序
echo.
echo 正在打包命令行程序...
pyinstaller --onefile --name="mail_exporter" --add-data="README.md;." --add-data="LICENSE;." mail_exporter.py
if errorlevel 1 (
    echo 错误: 命令行程序打包失败
    pause
    exit /b 1
)

:: 检查Inno Setup是否安装
echo.
echo 检查Inno Setup编译器...
set "INNO_SETUP_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%INNO_SETUP_PATH%" (
    set "INNO_SETUP_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"
)
if not exist "%INNO_SETUP_PATH%" (
    echo 警告: 未找到Inno Setup编译器
    echo 请从 https://jrsoftware.org/isinfo.php 下载并安装Inno Setup
    echo 或手动运行setup.iss文件
    echo.
    echo 打包的可执行文件位于 dist\ 目录中
    pause
    exit /b 0
)

:: 使用Inno Setup编译安装程序
echo 正在编译安装程序...
"%INNO_SETUP_PATH%" "setup.iss"
if errorlevel 1 (
    echo 错误: 安装程序编译失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo 可执行文件位于: dist\
echo 安装程序位于: dist\
echo.
echo 文件列表:
dir /b dist\*.exe
echo.
pause