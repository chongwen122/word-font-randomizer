# 字符级字体随机替换工具

一个强大的Word文档处理工具，能够将文档中的每个字符随机替换为不同的字体。

✨ 核心特性
🔤 字符级随机替换
每个字符独立处理 - 文档中的每个字符都会独立随机选择字体
真正的随机效果 - 同一个单词中的不同字母可能使用不同字体
保持可读性 - 智能算法确保替换后的文档仍然清晰可读

🖋️ 高级手写模拟效果
智能倾斜趋势模拟 - 模拟真实手写的自然倾斜变化
自然的手写纠正模式 - 每10-20个字符自动纠正倾斜趋势
可调节的模拟强度 - 5档强度调节，从轻微到真实手写效果
微颤效果 - 模拟手部微颤的自然波动

📊 多种随机化效果
随机行间距 - 每行之间的随机间距，可调节力度
随机字符大小 - 字符级字号随机化，限制相邻字符高度落差
随机行首缩进 - 每行前随机添加1-5个空格，可调节力度
界面滚动支持 - 使用鼠标滚轮滚动界面，提升用户体验

🎯 智能字体匹配
字符可用性检查 - 自动检测字体是否包含目标字符
智能回退机制 - 如果首选字体不包含字符，自动寻找替代字体
字体负载均衡 - 在所有可用字体中均匀分配字符

📊 详细统计分析
实时进度显示 - 处理过程中显示当前进
字体使用统计 - 显示每种字体的使用频率
字符处理报告 - 生成详细的转换报告

🎨 完美格式保留
格式无损 - 保留所有原始格式（粗体、斜体、颜色、大小等）
布局不变 - 保持段落、表格、页眉页脚的原始布局
兼容性强 - 支持复杂的Word文档格式

🖥️ 用户友好界面
直观操作 - 简单的点击式操作流程
实时反馈 - 处理过程中的实时状态更新
错误处理 - 友好的错误提示和解决方案

🚀 快速开始
系统要求
Python: 3.7 或更高版本
操作系统: Windows 10+, macOS 10.14+, Ubuntu 18.04+
内存: 至少 4GB RAM
存储: 100MB 可用空间

安装步骤
1. 下载项目
# 使用 Git 克隆
git clone https://github.com/chongwen122/word-font-randomizer.git
cd font-randomizer-project
# 或直接下载 ZIP 文件并解压

2. 安装依赖
pip install -r requirements.txt

3. 准备字体文件
# 创建字体文件夹（如果不存在）
mkdir fonts
# 将您的 .ttf 或 .otf 字体文件放入 fonts 文件夹
# 建议使用 5-10 个不同的字体文件以获得最佳效果

4. 启动程序
# Windows 用户
run.bat

# Linux/Mac 用户
./run.sh

# 或直接运行
python src/enhanced_font_randomizer.py

📖 使用指南
基本使用流程
1.启动程序
    运行启动脚本或直接执行Python文件
    程序会自动加载所有可用字体

2.选择输入文件
    点击"浏览"按钮选择要处理的Word文档
    支持所有标准.docx格式文件

3.设置输出路径
    程序会自动生成输出文件名
    可以手动修改保存位置和文件名

4.配置效果参数 (v1.4.0新功能)
    手写效果：启用手写模拟并调节自然度
    行间距随机：设置行间距随机力度
    字符大小随机：设置字号随机力度
    行首缩进：设置缩进随机力度

5.开始转换
    点击"开始字符级字体替换"按钮
    等待处理完成（大文档可能需要几分钟）

6.查看结果
    在输出位置查看转换后的文档
    检查程序日志了解详细统计信息

高级功能
# 字体管理
    字体预览 - 在界面中查看所有可用字体
    字体刷新 - 动态加载新添加的字体文件
    字符覆盖检查 - 确保字体支持目标字符集

# 处理选项
    字符级随机化 - 每个字符独立选择字体
    智能匹配 - 自动寻找包含字符的字体
    格式保留 - 完整保留原始文档格式

# 统计报告
    字符计数 - 显示处理的字符总数
    字体使用 - 每种字体的使用频率
    成功率 - 成功应用字体的字符比例

v1.4.0 新功能详解
🖋️ 智能手写模拟效果
倾斜趋势模拟：模拟真实书写时的自然倾斜变化
自动纠正机制：每10-20个字符自动回归基线，模拟书写纠正
强度调节：5档强度可选，从轻微倾斜到真实手写效果
微颤效果：添加随机微颤，使效果更加自然

📏 随机行间距
力度调节：5档力度调节，控制行间距变化范围
保持可读性：在保持文档可读性的前提下增加随机性
段落独立：每个段落独立设置行间距

🔠 随机字符大小
相邻限制：限制相邻字符的高度落差不超过0.5磅
力度控制：调节字号随机变化范围
基础保持：在基础字号±随机范围内变化

📐 随机行首缩进
空格随机：每行前随机添加1-5个空格
力度调节：控制缩进空格数量的变化范围
自然效果：模拟真实书写的不规则缩进

🛠️ 技术架构
# 核心组件
├── FontManager              # 字体管理和字符映射
│   ├── load_fonts()         # 加载字体文件
│   ├── get_font_for_char()  # 字符到字体映射
│   └── font_cache          # 字体信息缓存
├── FontRandomizerApp        # 主应用程序
│   ├── create_widgets()     # 界面创建
│   ├── convert_document()   # 文档处理
│   └── multi_threading     # 多线程处理
├── HandwritingSimulator     # 手写效果模拟 (v1.4.0新增)
│   ├── get_char_tilt()      # 计算字符倾斜
│   ├── _start_new_trend()   # 开始新趋势
│   └── _ease_in_out()       # 缓动函数
├── LineSpacingManager       # 行间距管理 (v1.4.0新增)
│   └── get_random_line_spacing() # 获取随机行间距
└── DocumentProcessor       # 文档处理引擎
    ├── process_paragraphs() # 段落处理
    ├── process_tables()     # 表格处理
    └── preserve_format()    # 格式保留

# 关键技术
    python-docx: Word文档读写和处理
    fontTools: 字体文件解析和字符集检查
    Tkinter: 跨平台图形用户界面
    多线程: 避免界面卡顿，提升用户体验

# 算法原理
1.字体预加载
（预加载所有字体的字符集信息）
for font in fonts:
    chars = extract_character_set(font)
    cache[font_name] = chars

2.字符处理
（为每个字符寻找合适字体）
for char in text:
    font = find_font_with_char(char)
    if font:
        apply_font(char, font)
    else:
        keep_original_font(char)

3.格式保留
（保存并恢复原始格式）
original_format = save_format(run)
new_run = create_new_run(char)
restore_format(new_run, original_format)

📁 项目结构
font-randomizer-project/
├── .gitignore                 # Git忽略规则
├── LICENSE                    # MIT许可证
├── README.md                  # 项目说明
├── requirements.txt           # Python依赖
├── run.bat                   # Windows启动脚本
├── run.sh                    # Linux/Mac启动脚本
├── setup.py                  # 安装脚本
├── fonts/                    # 字体目录
│   └── README_fonts.md       # 字体说明
├── src/                      # 源代码
│   ├── __init__.py
│   ├── enhanced_font_randomizer.py  # 主程序
│   └── font_manager.py       # 字体管理器
├── docs/                     # 文档
│   └── images/               # 截图资源
└── examples/                 # 示例文件
    ├── input_sample.docx     # 输入示例
    └── output_sample.docx    # 输出示例

🔧 故障排除
常见问题
Q: 程序无法启动
A: 确保已安装Python 3.7+，并正确安装依赖包
   运行: pip install -r requirements.txt

Q: 没有找到字体文件
A: 将.ttf或.otf文件放入fonts文件夹
   点击"刷新字体列表"或重启程序

Q: 转换后字符显示为方框
A: 当前字体不包含该字符，尝试添加更多字体文件
   程序会自动寻找包含该字符的字体

Q: 处理大文档时速度很慢
A: 这是正常现象，字符级处理需要为每个字符单独操作
   请耐心等待，程序会显示实时进度

Q: 程序在处理过程中崩溃
A: 可能是内存不足，尝试关闭其他程序
   或使用较小尺寸的文档进行测试

# 错误代码参考
错误代码	说明	解决方案
FONT_LOAD_ERROR	字体加载失败	检查字体文件完整性
DOCUMENT_READ_ERROR	文档读取失败	检查文档是否损坏
NO_FONTS_AVAILABLE	没有可用字体	添加字体文件到fonts文件夹
OUTPUT_WRITE_ERROR	输出写入失败	检查文件权限和路径

🤝 贡献指南
我们欢迎各种形式的贡献！

报告问题
使用 GitHub Issues 报告bug
提供详细的重现步骤和环境信息
包括错误日志和截图

功能请求
在Issues中描述您想要的功能
说明使用场景和预期效果
讨论实现方案和可行性

代码贡献
1.Fork 本仓库
2.创建功能分支 (git checkout -b feature/AmazingFeature)
3.提交更改 (git commit -m 'Add some AmazingFeature')
4.推送到分支 (git push origin feature/AmazingFeature)
5.开启Pull Request

开发环境设置
# 克隆仓库
git clone https://github.com/chongwen122/word-font-randomizer.git
cd word-font-randomizer

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Linux/Mac
# venv\Scripts\activate   # Windows

# 安装开发依赖
pip install -r requirements.txt

📄 许可证
本项目采用 MIT 许可证 - 详见 LICENSE 文件。

🙏 致谢
感谢以下开源项目的支持：
python-docx - Word文档处理
fontTools - 字体文件操作
Tkinter - 图形用户界面

📞 技术支持
文档: 查看 docs 目录获取详细文档
问题: 通过 GitHub Issues 报告问题
讨论: 加入我们的 Discussions
邮件: 发送邮件至 2086177709@qq.com

🚀 更新计划
即将到来
批量处理模式
字体预览功能
自定义替换规则
导出统计报告

未来规划
支持更多文档格式
云端字体库
插件系统
命令行工具版本

# 让每个字符都拥有独特的字体风格，为您的文档增添个性与创意！ ✍️🎨
# 如果您觉得这个项目有用，请给它一个 ⭐ 星标支持！