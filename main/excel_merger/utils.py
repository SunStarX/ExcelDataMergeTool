"""通用工具函数"""
import os
import re
import warnings

class Utils:
    """通用工具类"""
    
    @staticmethod
    def normalize_header(header):
        """
        标准化表头，用于不同文件表头的对齐
        
        参数:
            header: 原始表头文本
            
        返回:
            标准化后的表头文本
        """
        if not header:
            return ""
            
        # 转换为字符串并去除前后空格
        normalized = str(header).strip()
        
        # 去除特殊字符
        normalized = re.sub(r'[^\w\s]', '', normalized)
        
        # 转换为小写
        normalized = normalized.lower()
        
        # 替换空格为下划线
        normalized = normalized.replace(' ', '_')
        
        return normalized
    
    @staticmethod
    def is_amount_column(header):
        """
        判断是否为金额相关列
        
        参数:
            header: 表头文本
            
        返回:
            如果是金额相关列则返回True，否则返回False
        """
        from .config import Config
        
        if not header:
            return False
            
        # 转换为小写
        header_lower = header.lower()
        
        # 检查是否包含金额相关关键词
        for keyword in Config.AMOUNT_KEYWORDS:
            if keyword in header_lower:
                return True
                
        return False
    
    @staticmethod
    def ensure_dir_exists(path):
        """确保目录存在，如果不存在则创建"""
        if not os.path.exists(path):
            os.makedirs(path)
    
    @staticmethod
    def filter_warnings():
        """过滤不必要的警告信息"""
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')
