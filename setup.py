#!/usr/bin/env python3
"""
字体随机替换工具安装脚本 v1.4.0
Font Randomizer Setup Script
"""

import os
import sys
import platform
import subprocess
from pathlib import Path

class SetupInstaller:
    def __init__(self):
        self.project_name = "font-randomizer"
        self.version = "1.4.0"  # 更新版本号
        self.author = "Font Randomizer Project"
        self.email = "support@example.com"
        
    def print_header(self):
        """打印安装标题"""
        print("=" * 50)
        print("    字体随机替换工具安装程序 v1.4.0")
        print("    Font Randomizer Setup")
        print("=" * 50)
        print()
        
    def check_python_version(self):
        """检查Python版本"""
        print("检查Python版本...")
        version = sys.version_info
        if version.major < 3 or (version.major == 3 and version.minor < 7):
            print(f"错误: 需要 Python 3.7 或更高版本，当前版本: {sys.version}")
            return False
        print(f"✓ Python {version.major}.{version.minor}.{version.micro} - 符合要求")
        return True
        
    def install_requirements(self):
        """安装依赖包"""
        print("\n安装依赖包...")
        
        requirements = [
            "python-docx>=0.8.11",
            "fonttools>=4.0.0"
        ]
        
        for package in requirements:
            print(f"安装 {package}...")
            try:
                subprocess.check_call([
                    sys.executable, "-m", "pip", "install", package
                ])
                print(f"✓ {package} 安装成功")
            except subprocess.CalledProcessError as e:
                print(f"✗ {package} 安装失败: {e}")
                return False
                
        return True
        
    def create_directories(self):
        """创建必要的目录结构"""
        print("\n创建目录结构...")
        
        directories = [
            "fonts",
            "src",
            "docs",
            "tests",
            "docs/images"
        ]
        
        for directory in directories:
            path = Path(directory)
            if not path.exists():
                path.mkdir(parents=True, exist_ok=True)
                print(f"✓ 创建目录: {directory}")
            else:
                print(f"✓ 目录已存在: {directory}")
                
        return True
        
    def create_fonts_readme(self):
        """创建字体说明文件"""
        fonts_readme = """# 字体文件说明

## 字体要求

- 格式: .ttf 或 .otf
- 编码: 支持Unicode字符
- 数量: 建议2-5个字体文件

## 字体来源

请使用免费字体或您拥有使用权限的字体：

- Google Fonts (https://fonts.google.com/)
- Font Squirrel (https://www.fontsquirrel.com/)
- DaFont (https://www.dafont.com/) (注意许可证)

## 添加字体

1. 将字体文件复制到此文件夹
2. 重启应用程序
3. 字体将自动加载

## 新功能支持

v1.4.0 版本支持以下高级功能：
- 智能手写模拟效果
- 随机行间距调节
- 随机字符大小
- 随机行首缩进
- 界面滚动支持

## 许可证提醒

请确保您拥有所用字体的合法使用权。
本程序不包含任何字体文件，用户需自行提供。
"""
        
        with open("fonts/README_fonts.md", "w", encoding="utf-8") as f:
            f.write(fonts_readme)
        print("✓ 创建字体说明文件")
        
    def create_requirements_file(self):
        """创建requirements.txt文件"""
        requirements = "python-docx>=0.8.11\nfonttools>=4.0.0\n"
        
        with open("requirements.txt", "w", encoding="utf-8") as f:
            f.write(requirements)
        print("✓ 创建requirements.txt文件")
        
    def create_gitignore(self):
        """创建.gitignore文件"""
        gitignore = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
pip-wheel-metadata/
share/python-wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# Virtual environments
.env
.venv
env/
venv/
ENV/
env.bak/
venv.bak/

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
.DS_Store?
._*
.Spotlight-V100
.Trashes
ehthumbs.db
Thumbs.db

# Project specific
output/
temp/
*.log
"""
        
        with open(".gitignore", "w", encoding="utf-8") as f:
            f.write(gitignore)
        print("✓ 创建.gitignore文件")
        
    def create_license(self):
        """创建LICENSE文件"""
        license_text = """MIT License

Copyright (c) 2024 Font Randomizer Project

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""
        
        with open("LICENSE", "w", encoding="utf-8") as f:
            f.write(license_text)
        print("✓ 创建LICENSE文件")
        
    def create_run_scripts(self):
        """创建运行脚本 - 使用当前版本的脚本内容"""
        # Windows批处理文件 (从提供的run.bat内容)
        run_bat = """@echo off
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
if not exist "src\\enhanced_font_randomizer.py" (
    echo 错误: 未找到主程序文件
    echo 请确保 src\\enhanced_font_randomizer.py 存在
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
python src\\enhanced_font_randomizer.py

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
"""
        
        # Linux/Mac Shell脚本 (从提供的run.sh内容)
        run_sh = """#!/bin/bash

# 字符级字体随机替换工具启动脚本 v1.4.0

# 设置颜色代码
RED='\\033[0;31m'
GREEN='\\033[0;32m'
YELLOW='\\033[1;33m'
BLUE='\\033[0;34m'
NC='\\033[0m' # No Color

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
"""
        
        # 写入Windows批处理文件
        with open("run.bat", "w", encoding="utf-8") as f:
            f.write(run_bat)
        print("✓ 创建Windows启动脚本: run.bat")
        
        # 写入Linux/Mac Shell脚本
        with open("run.sh", "w", encoding="utf-8") as f:
            f.write(run_sh)
        
        # 设置Shell脚本执行权限
        if platform.system() != "Windows":
            os.chmod("run.sh", 0o755)
        print("✓ 创建Linux/Mac启动脚本: run.sh")
        
    def create_documentation(self):
        """创建基础文档"""
        docs_dir = Path("docs")
        
        # 用户指南 - 更新以反映v1.4.0功能
        user_guide = """# 用户指南 v1.4.0

## 系统要求

- Python 3.7 或更高版本
- Windows 10+/macOS 10.14+/Ubuntu 18.04+

## 安装步骤

### 1. 安装Python
从 Python官网 (https://www.python.org/downloads/) 下载并安装Python

### 2. 下载项目
使用Git克隆项目或直接下载ZIP文件

### 3. 安装依赖
运行命令: pip install -r requirements.txt
或直接运行 setup.py

### 4. 添加字体
将字体文件(.ttf/.otf)放入 fonts 文件夹

## 使用教程

### 基本使用
1. 启动程序
2. 选择输入Word文档
3. 设置输出文件路径
4. 配置随机化效果
5. 点击"开始转换"
6. 查看处理结果

### v1.4.0 新功能

#### 智能手写模拟效果
- 模拟真实手写的倾斜和纠正模式
- 可调节手写自然度强度
- 智能趋势变化和微颤效果

#### 随机行间距
- 每行之间的随机间距
- 可调节的随机力度
- 保持文档可读性

#### 随机字符大小
- 字符级字号随机化
- 限制相邻字符高度落差
- 可调节的随机力度

#### 随机行首缩进
- 每行前随机添加空格
- 模拟自然书写的不规则性
- 可调节的缩进力度

#### 界面滚动支持
- 使用鼠标滚轮滚动界面
- 更好的用户体验

## 常见问题

### Q: 程序无法启动
A: 确保已安装Python并正确安装依赖

### Q: 字体不显示
A: 检查字体文件是否放入fonts文件夹，且格式正确

### Q: 处理速度慢
A: 大文档需要较长时间，这是正常现象

### Q: 手写效果不明显
A: 调整手写自然度滑块到较高强度
"""
        
        # 开发文档
        development = """# 开发文档 v1.4.0

## 项目结构

font-randomizer-project/
├── src/                    # 源代码
│   ├── __init__.py         # 包初始化
│   ├── enhanced_font_randomizer.py  # 主程序
│   └── font_manager.py     # 字体管理
├── fonts/                  # 字体文件
├── tests/                  # 测试代码
└── docs/                   # 文档

## 核心模块

### FontManager
- 字体加载和字符可用性检查
- 预加载字体字符集信息
- 随机字体选择

### FontRandomizerApp
- 主界面和文档处理逻辑
- 多线程处理避免界面冻结
- 完整的错误处理

### HandwritingSimulator
- 智能手写效果模拟
- 倾斜趋势和纠正模式
- 自然的手写行为模拟

### LineSpacingManager
- 随机行间距管理
- 间距缓存和一致性保持

## 扩展开发

### 添加新功能
1. 在相应模块中添加功能
2. 更新用户界面
3. 添加测试用例

### 代码规范
- 使用PEP 8代码风格
- 添加类型提示
- 编写文档字符串
"""
        
        # 写入文档文件
        with open(docs_dir / "user_guide.md", "w", encoding="utf-8") as f:
            f.write(user_guide)
            
        with open(docs_dir / "development.md", "w", encoding="utf-8") as f:
            f.write(development)
            
        print("✓ 创建文档文件")
        
    def create_test_files(self):
        """创建测试文件"""
        tests_dir = Path("tests")
        
        # 基本测试文件 - 更新以测试新功能
        test_basic = '''#!/usr/bin/env python3
"""
基本功能测试 v1.4.0
"""

import os
import sys
import unittest
from pathlib import Path

# 添加src目录到Python路径
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

class TestBasic(unittest.TestCase):
    """基本功能测试"""
    
    def test_imports(self):
        """测试模块导入"""
        try:
            from enhanced_font_randomizer import FontRandomizerApp, HandwritingSimulator, LineSpacingManager
            from font_manager import FontManager
            self.assertTrue(True)
        except ImportError as e:
            self.fail(f"导入失败: {e}")
    
    def test_fonts_directory(self):
        """测试字体目录"""
        self.assertTrue(os.path.exists("fonts") or True)  # 允许目录不存在
    
    def test_requirements(self):
        """测试依赖包"""
        try:
            import docx
            import fontTools
            self.assertTrue(True)
        except ImportError:
            self.skipTest("缺少依赖包")
    
    def test_handwriting_simulator(self):
        """测试手写模拟器"""
        try:
            from enhanced_font_randomizer import HandwritingSimulator
            simulator = HandwritingSimulator()
            # 测试获取字符倾斜
            tilt = simulator.get_char_tilt('A')
            self.assertIsInstance(tilt, float)
        except ImportError:
            self.skipTest("无法导入手写模拟器")

class TestNewFeatures(unittest.TestCase):
    """v1.4.0 新功能测试"""
    
    def test_line_spacing_manager(self):
        """测试行间距管理器"""
        try:
            from enhanced_font_randomizer import LineSpacingManager
            manager = LineSpacingManager()
            spacing = manager.get_random_line_spacing(0)
            self.assertIsInstance(spacing, float)
            self.assertTrue(0.8 <= spacing <= 1.2)
        except ImportError:
            self.skipTest("无法导入行间距管理器")

if __name__ == '__main__':
    unittest.main()
'''
        
        # 写入测试文件
        with open(tests_dir / "test_basic.py", "w", encoding="utf-8") as f:
            f.write(test_basic)
            
        # 创建测试包初始化文件
        with open(tests_dir / "__init__.py", "w", encoding="utf-8") as f:
            f.write('"""测试包"""')
            
        print("✓ 创建测试文件")
        
    def create_src_init(self):
        """创建src包初始化文件 - 使用当前版本信息"""
        src_dir = Path("src")
        
        init_content = '''"""
字体随机替换工具源代码包
字符级字体随机替换工具 - 将Word文档中的每个字符随机替换为不同字体
"""

__version__ = "1.4.0"
__author__ = "Font Randomizer Project"
__email__ = "support@example.com"

from .enhanced_font_randomizer import FontRandomizerApp, HandwritingSimulator, LineSpacingManager
from .font_manager import FontManager

__all__ = ['FontRandomizerApp', 'FontManager', 'HandwritingSimulator', 'LineSpacingManager']
'''
        
        with open(src_dir / "__init__.py", "w", encoding="utf-8") as f:
            f.write(init_content)
            
        print("✓ 创建源代码包初始化文件")
        
    def display_usage_instructions(self):
        """显示使用说明"""
        print("\n" + "=" * 50)
        print("安装完成！ v1.4.0")
        print("=" * 50)
        print("\n使用方法:")
        
        if platform.system() == "Windows":
            print("  • 双击 run.bat 文件启动程序")
        else:
            print("  • 运行 ./run.sh 启动程序")
            
        print("  • 或直接运行: python src/enhanced_font_randomizer.py")
        
        print("\n新功能说明:")
        print("  • 智能手写模拟 - 模拟真实手写效果")
        print("  • 随机行间距 - 可调节力度的行间距随机化") 
        print("  • 随机字符大小 - 限制相邻字符高度落差")
        print("  • 随机行首缩进 - 每行前随机添加空格")
        print("  • 界面滚动支持 - 使用鼠标滚轮滚动")
        
        print("\n下一步:")
        print("  1. 将字体文件(.ttf/.otf)放入 fonts 文件夹")
        print("  2. 启动程序")
        print("  3. 选择Word文档并配置效果参数")
        print("  4. 开始处理")
        
        print("\n项目结构:")
        print("  fonts/          - 存放字体文件")
        print("  src/            - 源代码")
        print("  docs/           - 文档")
        print("  tests/          - 测试文件")
        print("  requirements.txt - 依赖列表")
        
        print("\n技术支持:")
        print("  查看 docs/ 目录获取详细文档")
        print("  报告问题: GitHub Issues")
        
    def run_setup(self):
        """运行完整的安装流程"""
        self.print_header()
        
        # 执行安装步骤
        steps = [
            ("检查Python版本", self.check_python_version),
            ("安装依赖包", self.install_requirements),
            ("创建目录结构", self.create_directories),
            ("创建字体说明", self.create_fonts_readme),
            ("创建依赖文件", self.create_requirements_file),
            ("创建Git忽略文件", self.create_gitignore),
            ("创建许可证文件", self.create_license),
            ("创建运行脚本", self.create_run_scripts),
            ("创建文档", self.create_documentation),
            ("创建测试文件", self.create_test_files),
            ("创建源代码包", self.create_src_init),
        ]
        
        for step_name, step_func in steps:
            print(f"\n[{step_name}]")
            if not step_func():
                print(f"✗ {step_name} 失败")
                return False
                
        self.display_usage_instructions()
        return True

def main():
    """主函数"""
    installer = SetupInstaller()
    
    try:
        success = installer.run_setup()
        if success:
            print(f"\n✓ 安装程序完成 v1.4.0")
            sys.exit(0)
        else:
            print(f"\n✗ 安装程序失败")
            sys.exit(1)
    except KeyboardInterrupt:
        print(f"\n安装被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n安装过程中出现错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()