import os
import random
import glob
from fontTools.ttLib import TTFont

class FontManager:
    """
    字体管理器
    负责加载字体文件、检查字符可用性和字体选择
    """
    
    def __init__(self, fonts_dir="fonts"):
        self.fonts_dir = fonts_dir
        self.font_files = []
        self.font_cache = {}  # 字体名称 -> {path, chars, object}
        self.load_fonts()
    
    def load_fonts(self):
        """
        加载所有字体文件并预加载字符信息
        
        Returns:
            int: 成功加载的字体数量
        """
        self.font_files = []
        self.font_cache = {}
        
        if not os.path.exists(self.fonts_dir):
            print(f"字体目录不存在: {self.fonts_dir}")
            return 0
        
        loaded_count = 0
        for ext in ['*.ttf', '*.otf']:
            font_paths = glob.glob(os.path.join(self.fonts_dir, ext))
            for font_path in font_paths:
                try:
                    # 加载字体文件
                    font = TTFont(font_path)
                    font_name = os.path.splitext(os.path.basename(font_path))[0]
                    
                    # 获取字体支持的字符集
                    chars = set()
                    for table in font['cmap'].tables:
                        chars.update(table.cmap.keys())
                    
                    # 缓存字体信息
                    self.font_cache[font_name] = {
                        'path': font_path,
                        'chars': chars,
                        'object': font
                    }
                    self.font_files.append(font_path)
                    loaded_count += 1
                    
                    print(f"加载字体: {font_name} (包含 {len(chars)} 个字符)")
                    
                except Exception as e:
                    print(f"加载字体 {font_path} 时出错: {e}")
        
        return loaded_count
    
    def get_font_for_char(self, char):
        """
        为指定字符查找可用的字体
        
        Args:
            char (str): 要查找字体的字符
            
        Returns:
            str or None: 字体名称，如果找不到返回None
        """
        if not self.font_cache:
            return None
            
        char_code = ord(char)
        
        # 收集所有支持该字符的字体
        supported_fonts = []
        for font_name, font_info in self.font_cache.items():
            if char_code in font_info['chars']:
                supported_fonts.append(font_name)
        
        # 在支持字符的字体中随机选择
        if supported_fonts:
            return random.choice(supported_fonts)
        else:
            # 如果没有字体支持该字符，返回None
            return None
    
    def get_random_font_name(self):
        """
        获取随机字体名称
        
        Returns:
            str or None: 随机字体名称
        """
        if not self.font_cache:
            return None
        return random.choice(list(self.font_cache.keys()))
    
    def get_font_count(self):
        """获取字体数量"""
        return len(self.font_cache)
    
    def get_font_names(self):
        """获取所有字体名称"""
        return list(self.font_cache.keys())
    
    def get_font_info(self, font_name):
        """
        获取字体详细信息
        
        Args:
            font_name (str): 字体名称
            
        Returns:
            dict or None: 字体信息
        """
        return self.font_cache.get(font_name)
    
    def is_char_supported(self, font_name, char):
        """
        检查字体是否支持指定字符
        
        Args:
            font_name (str): 字体名称
            char (str): 要检查的字符
            
        Returns:
            bool: 是否支持
        """
        font_info = self.font_cache.get(font_name)
        if not font_info:
            return False
        
        char_code = ord(char)
        return char_code in font_info['chars']