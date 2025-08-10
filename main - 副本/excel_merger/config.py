"""配置常量定义"""

class Config:
    """配置常量定义"""
    EXCEL_EXTENSIONS = ('.xlsx', '.xls')
    TEMP_FILE_PREFIX = '~$'
    AMOUNT_KEYWORDS = ["金额", "结算", "价格", "费用", "总价", "合计", "付款", "收款"]
    EMPTY_COLUMN_THRESHOLD = 5  # 连续空列判断表头结束的阈值
    LONG_NUMBER_THRESHOLD = 11  # 长数字判断阈值
    DEFAULT_OUTPUT_DIR = "."  # 默认输出目录
