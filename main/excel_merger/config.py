"""配置常量定义 - 增加最大列数限制"""
class Config:
    """应用程序配置常量"""
    # Excel文件扩展名
    EXCEL_EXTENSIONS = ('.xlsx', '.xlsm')

    # 临时文件前缀（会被忽略）
    TEMP_FILE_PREFIX = '~$'

    # 判断表头结束的连续空列阈值
    EMPTY_COLUMN_THRESHOLD = 3

    # 长数字阈值（超过此长度将使用文本格式）
    LONG_NUMBER_THRESHOLD = 10

    # 金额列关键词（用于识别金额相关列）
    AMOUNT_KEYWORDS = {'金额', '钱', '款', '费用', '总计', '合计', 'sum', 'amount'}

    # 最大检查列数（防止无限循环）
    MAX_COLUMNS_TO_CHECK = 100  # 合理的列数限制

    # 写入批处理大小（提高大文件处理效率）
    WRITE_BATCH_SIZE = 100
