"""程序主控制器"""

from .excel_processor import ExcelProcessor
from .utils import Utils

class ExcelMerger:
    """Excel文件合并工具主控制器"""
    def __init__(self, folder_path: str = '.', output_dir: str = '.'):
        self.processor = ExcelProcessor(folder_path, output_dir)

    def merge(self) -> None:
        """执行合并操作的主方法"""
        try:
            # 初始化和准备
            Utils.filter_warnings()
            self.processor.get_excel_files()
            
            # 分析第一个文件
            first_file = self.processor.excel_files[0]
            first_row_count, start_row, format_ref_row = self.processor.analyze_first_file(first_file)
            
            # 处理其他文件
            merged_df = self.processor.process_other_files(first_file)
            
            # 写入合并数据
            self.processor.write_merged_data(merged_df, first_row_count, start_row, format_ref_row)
            
            # 保存结果
            self.processor.save_result()
            
        except Exception as e:
            print(f"\n操作失败: {str(e)}")

def main():
    """主函数入口"""
    print("===== Excel文件合并工具 =====")
    # 可在此处指定文件夹路径和输出目录
    # merger = ExcelMerger("path/to/excel/files", "path/to/output")
    merger = ExcelMerger()
    merger.merge()
    input("\n按回车键退出...")
