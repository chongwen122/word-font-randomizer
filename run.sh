#!/bin/bash

# 字符级字体随机替换工具启动脚本 v1.4.0

# 设置颜色代码
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 显示标题
echo -e "${BLUE}"
echo "========================================"
echo "    字符级字体随机替换工具 v1.4.0"
echo "    Character Level Font Randomizer"
echo "========================================"
echo -e "${NC}"

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}错误: 未找到Python3${NC}"
    echo "请先安装Python 3.7或更高版本"
    echo "下载地址: https://www.python.org/downloads/"
    echo ""
    exit 1
fi

# 显示Python版本
echo -e "${GREEN}检测到Python版本:${NC}"
python3 --version
echo ""

# 检查依赖包
echo -e "${YELLOW}检查依赖包...${NC}"

# 检查python-docx
if ! python3 -c "import docx" &> /dev/null; then
    echo -e "${YELLOW}安装 python-docx...${NC}"
    pip3 install python-docx
    if [ $? -ne 0 ]; then
        echo -e "${RED}安装 python-docx 失败${NC}"
        echo "请尝试手动安装: pip3 install python-docx"
        exit 1
    fi
else
    echo -e "${GREEN}python-docx 已安装${NC}"
fi

# 检查fonttools
if ! python3 -c "import fontTools" &> /dev/null; then
    echo -e "${YELLOW}安装 fonttools...${NC}"
    pip3 install fonttools
    if [ $? -ne 0 ]; then
        echo -e "${RED}安装 fonttools 失败${NC}"
        echo "请尝试手动安装: pip3 install fonttools"
        exit 1
    fi
else
    echo -e "${GREEN}fonttools 已安装${NC}"
fi

# 检查字体目录
if [ ! -d "fonts" ]; then
    echo -e "${YELLOW}创建字体目录...${NC}"
    mkdir fonts
    echo -e "${YELLOW}请将 .ttf 或 .otf 字体文件放入 fonts 文件夹${NC}"
fi

# 检查源代码目录
if [ ! -d "src" ]; then
    echo -e "${RED}错误: 未找到 src 目录${NC}"
    echo "请确保程序文件完整"
    exit 1
fi

# 检查主程序文件
if [ ! -f "src/enhanced_font_randomizer.py" ]; then
    echo -e "${RED}错误: 未找到主程序文件${NC}"
    echo "请确保 src/enhanced_font_randomizer.py 存在"
    exit 1
fi

# 运行主程序
echo ""
echo -e "${GREEN}启动主程序...${NC}"
echo -e "${YELLOW}如果遇到问题，请确保:${NC}"
echo "1. 字体文件已放入 fonts 文件夹"
echo "2. 所有依赖包已正确安装"
echo "3. 输入的Word文档格式正确"
echo ""
echo -e "${YELLOW}新功能:${NC}"
echo -e "${BLUE}- 智能手写模拟效果 - 模拟真实手写的倾斜和纠正模式${NC}"
echo -e "${BLUE}- 随机行间距 - 可调节力度的行间距随机化${NC}"
echo -e "${BLUE}- 随机字符大小 - 可调节力度的字号随机化，限制相邻字符高度落差${NC}"
echo -e "${BLUE}- 随机行首缩进 - 可调节力度的行首空格随机化${NC}"
echo -e "${BLUE}- 界面滚动支持 - 可使用鼠标滚轮滚动界面${NC}"
echo ""
echo -e "${YELLOW}按 Ctrl+C 可随时退出程序${NC}"
echo "========================================"

# 运行Python程序
python3 src/enhanced_font_randomizer.py

# 检查程序退出状态
EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo -e "${RED}程序异常退出，代码: $EXIT_CODE${NC}"
else
    echo ""
    echo -e "${GREEN}程序已退出${NC}"
fi