@echo off
chcp 65001 > nul
title 字符级字体随机替换工具

echo ========================================
echo    字符级字体随机替换工具 v1.4.0
echo    Character Level Font Randomizer
echo ========================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python
    echo 请先安装Python 3.7或更高版本
    echo 下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: 显示Python版本
echo 检测到Python版本:
python --version
echo.

:: 检查依赖包
echo 检查依赖包...
pip show python-docx >nul 2>&1
if errorlevel 1 (
    echo 安装 python-docx...
    pip install python-docx
) else (
    echo python-docx 已安装
)

pip show fonttools >nul 2>&1
if errorlevel 1 (
    echo 安装 fonttools...
    pip install fonttools
) else (
    echo fonttools 已安装
)

:: 检查字体目录
if not exist "fonts" (
    echo 创建字体目录...
    mkdir fonts
    echo 请将 .ttf 或 .otf 字体文件放入 fonts 文件夹
)

:: 检查源代码目录
if not exist "src" (
    echo 错误: 未找到 src 目录
    echo 请确保程序文件完整
    pause
    exit /b 1
)

:: 检查主程序文件
if not exist "src\enhanced_font_randomizer.py" (
    echo 错误: 未找到主程序文件
    echo 请确保 src\enhanced_font_randomizer.py 存在
    pause
    exit /b 1
)

:: 运行主程序
echo.
echo 启动主程序...
echo 如果遇到问题，请确保:
echo 1. 字体文件已放入 fonts 文件夹
echo 2. 所有依赖包已正确安装
echo 3. 输入的Word文档格式正确
echo.
echo 新功能:
echo - 智能手写模拟效果 - 模拟真实手写的倾斜和纠正模式
echo - 随机行间距 - 可调节力度的行间距随机化
echo - 随机字符大小 - 可调节力度的字号随机化，限制相邻字符高度落差
echo - 随机行首缩进 - 可调节力度的行首空格随机化
echo - 界面滚动支持 - 可使用鼠标滚轮滚动界面
echo.
echo 按 Ctrl+C 可随时退出程序
echo ========================================
python src\enhanced_font_randomizer.py

:: 如果程序正常退出，暂停以便查看输出
if errorlevel 1 (
    echo.
    echo 程序异常退出，代码: %errorlevel%
    pause
) else (
    echo.
    echo 程序已退出
    pause
)