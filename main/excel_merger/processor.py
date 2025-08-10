"""Excel处理核心逻辑 - 修复StyleProxy哈希错误"""
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from .config import Config
from .utils import Utils
from .format_handler import FormatHandler

class ExcelProcessor:
    """Excel文件处理核心类"""
    def __init__(self, folder_path, output_path, log_callback):
        self.folder_path = folder_path
        self.output_path = output_path
        self.log = log_callback  # 日志回调函数
        self.excel_files = []
        self.header_info = []  # (列索引, 原始表头, 标准化表头, 是否金额列)
        self.header_map = {}
        self.wb = None
        self.ws = None

        # 初始化时过滤警告
        Utils.filter_warnings()

    def merge(self):
        """执行合并操作"""
        # 获取Excel文件
        self._get_excel_files()

        # 分析第一个文件
        first_file = self.excel_files[0]
        first_row_count, start_row, format_ref_row = self._analyze_first_file(first_file)

        # 处理其他文件
        merged_df = self._process_other_files(first_file)

        # 写入合并数据
        self._write_merged_data(merged_df, first_row_count, start_row, format_ref_row)

        # 保存结果
        return self._save_result()

    def _get_excel_files(self):
        """获取文件夹中所有Excel文件"""
        self.log(f"正在扫描文件夹: {self.folder_path}")

        self.excel_files = [
            f for f in os.listdir(self.folder_path)
            if f.endswith(Config.EXCEL_EXTENSIONS) and not f.startswith(Config.TEMP_FILE_PREFIX)
        ]

        if not self.excel_files:
            raise FileNotFoundError("未找到任何Excel文件")

        self.log(f"找到{len(self.excel_files)}个Excel文件")

    def _analyze_first_file(self, first_file):
        """分析第一个文件，获取表头信息"""
        first_file_path = os.path.join(self.folder_path, first_file)
        self.log(f"以文件 '{first_file}' 为基础进行合并")

        # 加载第一个文件
        try:
            self.wb = load_workbook(first_file_path)
        except Exception as e:
            raise IOError(f"无法加载文件 {first_file}: {str(e)}")

        self.ws = self.wb.active

        # 分析表头（添加最大列数限制）
        col_idx = 1
        while col_idx <= Config.MAX_COLUMNS_TO_CHECK:
            original_header = self.ws.cell(row=1, column=col_idx).value
            normalized = Utils.normalize_header(original_header)
            is_amount_col = Utils.is_amount_column(str(original_header)) if original_header else False

            self.header_info.append((col_idx, original_header, normalized, is_amount_col))

            if normalized and normalized not in self.header_map:
                self.header_map[normalized] = col_idx

            # 检测表头结束
            if self._is_header_end(col_idx):
                break

            col_idx += 1

        # 显示检测到的金额列
        self._display_amount_columns()

        # 读取第一个文件数据
        try:
            first_df = pd.read_excel(first_file_path, dtype=str)
        except Exception as e:
            raise IOError(f"无法读取文件 {first_file}: {str(e)}")

        first_row_count = len(first_df)
        start_row = first_row_count + 2  # 数据开始追加的位置
        format_ref_row = 2 if first_row_count > 0 else 1

        return first_row_count, start_row, format_ref_row

    def _is_header_end(self, col_idx):
        """判断表头是否结束"""
        empty_count = 0
        for i in range(1, Config.EMPTY_COLUMN_THRESHOLD + 1):
            if col_idx + i > Config.MAX_COLUMNS_TO_CHECK:  # 防止超过最大列数
                return True
            if self.ws.cell(row=1, column=col_idx + i).value is None:
                empty_count += 1
        return empty_count >= Config.EMPTY_COLUMN_THRESHOLD

    def _display_amount_columns(self):
        """显示检测到的金额列"""
        self.log(f"基础文件表头分析完成: 共 {len(self.header_info)} 列")
        amount_columns = [
            f"第{idx}列: {orig}"
            for idx, orig, _, is_amt in self.header_info
            if is_amt
        ]
        if amount_columns:
            self.log(f"检测到金额相关列: {', '.join(amount_columns)}")

    def _process_other_files(self, first_file):
        """处理其他文件并合并数据"""
        try:
            all_data = [pd.read_excel(os.path.join(self.folder_path, first_file), dtype=str)]
        except Exception as e:
            raise IOError(f"无法读取基础文件 {first_file}: {str(e)}")

        for file in self.excel_files[1:]:
            file_path = os.path.join(self.folder_path, file)
            try:
                df = pd.read_excel(file_path, dtype=str)
                file_headers = [Utils.normalize_header(col) for col in df.columns]

                # 对齐列
                aligned_df = pd.DataFrame(columns=[h[2] for h in self.header_info])
                for norm_header in aligned_df.columns:
                    if norm_header in file_headers:
                        df_col_idx = file_headers.index(norm_header)
                        aligned_df[norm_header] = df.iloc[:, df_col_idx]
                    else:
                        aligned_df[norm_header] = ""

                all_data.append(aligned_df)
                self.log(f"已处理文件: {file}")

            except Exception as e:
                self.log(f"警告: 处理文件{file}时出错，已跳过 - {str(e)}")

        if len(all_data) <= 1:
            self.log("警告: 仅找到一个有效文件或所有其他文件处理失败")

        # 合并所有数据
        merged_df = pd.concat(all_data, ignore_index=True)
        self.log(f"数据合并完成，共 {len(merged_df)} 行数据")
        return merged_df

    def _write_merged_data(self, merged_df, first_row_count, start_row, format_ref_row):
        """将合并后的数据写入Excel"""
        # 清除原有多余数据
        if self.ws.max_row >= start_row:
            try:
                self.ws.delete_rows(start_row, self.ws.max_row - start_row + 1)
                self.log("已清除基础文件后的原有数据")
            except Exception as e:
                self.log(f"警告: 清除旧数据时出错 - {str(e)}")

        # 分批写入数据，提高大文件处理效率
        total_rows = len(merged_df) - first_row_count
        for batch_start in range(0, total_rows, Config.WRITE_BATCH_SIZE):
            batch_end = min(batch_start + Config.WRITE_BATCH_SIZE, total_rows)
            self.log(f"正在写入数据: {batch_end}/{total_rows} 行")

            for row_offset in range(batch_start, batch_end):
                row_idx = first_row_count + row_offset
                data_row = merged_df.iloc[row_idx]
                current_row = start_row + row_offset

                for col_info in self.header_info:
                    self._write_cell(data_row, col_info, current_row, format_ref_row)

    def _write_cell(self, data_row, col_info, current_row, format_ref_row):
        """写入单个单元格数据 - 修复StyleProxy哈希错误"""
        col_idx, orig_header, norm_header, is_amount_col = col_info

        # 获取单元格值
        value = data_row[norm_header] if norm_header in data_row else ""
        if pd.isna(value):
            value = ""

        # 获取参考单元格和目标单元格
        ref_cell = self.ws.cell(row=format_ref_row, column=col_idx)
        target_cell = self.ws.cell(row=current_row, column=col_idx)

        # 处理金额列
        if is_amount_col:
            value = self._process_amount_value(value, target_cell, ref_cell)
        # 处理长数字
        elif str(value).isdigit() and len(str(value)) > Config.LONG_NUMBER_THRESHOLD:
            target_cell.number_format = '@'
            value = str(value)
        else:
            # 直接复制格式属性而非整个StyleProxy对象
            target_cell.number_format = ref_cell.number_format

        target_cell.value = value

        # 复制格式时避免哈希操作
        FormatHandler.copy_cell_format(ref_cell, target_cell, force_right=is_amount_col)

    def _process_amount_value(self, value, target_cell, ref_cell):
        """处理金额列的值"""
        try:
            # 清除货币符号和千位分隔符
            clean_value = str(value).replace(',', '').replace('￥', '').replace('$', '')
            value = float(clean_value)
            target_cell.number_format = ref_cell.number_format
            return value
        except:
            # 转换失败时保持文本格式但强制右对齐
            target_cell.number_format = '@'
            return value

    def _save_result(self):
        """保存合并结果"""
        # 确保输出目录存在
        Utils.ensure_dir_exists(self.output_path)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(self.output_path, f"汇总结果_{timestamp}.xlsx")
        
        try:
            self.wb.save(output_file)
            return output_file
        except Exception as e:
            raise IOError(f"保存文件失败: {str(e)}")
