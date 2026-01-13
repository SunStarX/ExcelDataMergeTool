"""通用工具函数"""
import re
import os
import warnings
import pandas as pd  # 新增：导入pandas用于CSV文件读取

class Utils:
    """通用工具类"""

    @staticmethod
    def normalize_header(header):
        """
        标准化表头，用于不同文件间的表头匹配

        参数:
            header: 原始表头字符串

        返回:
            标准化后的表头字符串
        """
        if not header:
            return ""

        # 转换为字符串并去除首尾空格
        normalized = str(header).strip()

        # 移除所有特殊字符
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
            header: 表头字符串

        返回:
            如果是金额列则返回True，否则返回False
        """
        from .config import Config

        if not header:
            return False

        # 标准化表头
        normalized_header = Utils.normalize_header(header)

        # 检查是否包含金额关键词
        for keyword in Config.AMOUNT_KEYWORDS:
            if keyword in normalized_header:
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

    # 新增：判断文件是否为CSV格式
    @staticmethod
    def is_csv_file(file_path):
        """
        判断文件是否为CSV格式（仅通过后缀判断）

        参数:
            file_path: 文件完整路径

        返回:
            是CSV返回True，否则返回False
        """
        if not file_path:
            return False
        ext = os.path.splitext(file_path)[1].lower()
        return ext in ['.csv']

    # 新增：读取CSV文件（适配中文编码）
    @staticmethod
    def read_csv_file(file_path):
        """
        读取CSV文件为pandas.DataFrame

        参数:
            file_path: CSV文件路径

        返回:
            pandas.DataFrame | None（读取失败返回None）
        """
        try:
            # 优先utf-8编码，失败则尝试gbk（适配中文CSV文件）
            df = pd.read_csv(file_path, encoding='utf-8')
            return df
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='gbk')
                return df
            except Exception as e:
                warnings.warn(f"读取CSV文件失败 {file_path}: {str(e)}")
                return None
        except Exception as e:
            warnings.warn(f"读取CSV文件异常 {file_path}: {str(e)}")
            return None