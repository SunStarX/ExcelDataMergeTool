"""通用工具函数"""

import pandas as pd
import warnings

class Utils:
    """通用工具函数"""
    @staticmethod
    def normalize_header(header) -> str:
        """标准化表头，处理空格、大小写等差异"""
        if pd.isna(header):
            return ""
        normalized = str(header).strip().lower()
        normalized = normalized.replace(" ", "").replace("-", "").replace("_", "")
        normalized = normalized.replace("：", ":").replace("（", "(").replace("）", ")")
        return normalized

    @staticmethod
    def is_amount_column(header: str, keywords) -> bool:
        """判断是否为金额相关列"""
        header_lower = header.lower()
        return any(keyword in header_lower for keyword in keywords)

    @staticmethod
    def filter_warnings() -> None:
        """过滤不必要的警告"""
        warnings.filterwarnings(
            "ignore", 
            category=UserWarning, 
            message="Workbook contains no default style, apply openpyxl's default"
        )

    @staticmethod
    def ensure_dir_exists(path: str) -> None:
        """确保目录存在，不存在则创建"""
        import os
        if not os.path.exists(path):
            os.makedirs(path)
