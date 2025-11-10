"""
字体随机替换工具源代码包
字符级字体随机替换工具 - 将Word文档中的每个字符随机替换为不同字体
"""

__version__ = "1.4.0"
__author__ = "Font Randomizer Project"
__email__ = "support@example.com"

from .enhanced_font_randomizer import FontRandomizerApp, HandwritingSimulator, LineSpacingManager
from .font_manager import FontManager

__all__ = ['FontRandomizerApp', 'FontManager', 'HandwritingSimulator', 'LineSpacingManager']