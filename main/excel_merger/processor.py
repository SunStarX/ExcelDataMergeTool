"""Excel处理核心逻辑 - 按目录原生顺序合并文件"""
import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from .config import Config
from .utils import Utils
from .format_handler import FormatHandler

class ExcelProcessor:
    """Excel文件处理核心类 - 按目录原生顺序合并文件"""
    def __init__(self, folder_path, output_path, log_callback):
        self.folder_path = folder_path
        self.output_path = output_path
        self.log = log_callback  # 日志回调函数
        self.excel_files = []
        self.header_info = []  # (列索引, 原始表头, 标准化表头, 是否金额列)
        self.header_map = {}
        self.wb = None
        self.ws = None
        self.has_shop_column = False  # 是否已添加店铺列
        
        # 初始化时过滤警告
        Utils.filter_warnings()

    def merge(self):
        """执行合并操作 - 按目录原生顺序合并文件"""
        # 获取Excel文件（按目录原生顺序）
        self._get_excel_files_in_native_order()
        
        # 分析第一个文件获取表头信息
        first_file = self.excel_files[0]
        first_row_count, start_row, format_ref_row = self._analyze_first_file(first_file)
        
        # 添加店铺列到第一个文件
        self._add_shop_column_to_first_file(first_file, first_row_count)
        
        # 按原生顺序处理所有文件
        merged_df = self._process_all_files_in_native_order()
        
        # 写入合并数据
        self._write_merged_data(merged_df, first_row_count, start_row, format_ref_row)
        
        # 保存结果
        return self._save_result()

    def _get_excel_files_in_native_order(self):
        """获取文件夹中所有Excel文件，保持操作系统原生顺序"""
        self.log(f"正在扫描文件夹: {self.folder_path}")
        
        # 获取所有Excel文件，不进行排序，保持操作系统返回的原生顺序
        self.excel_files = [
            f for f in os.listdir(self.folder_path) 
            if f.endswith(Config.EXCEL_EXTENSIONS) and not f.startswith(Config.TEMP_FILE_PREFIX)
        ]
        
        if not self.excel_files:
            raise FileNotFoundError("未找到任何Excel文件")
            
        self.log(f"找到{len(self.excel_files)}个Excel文件，将按以下原生顺序处理:")
        for i, file in enumerate(self.excel_files, 1):
            self.log(f"  {i}. {file}")

    def _add_shop_column_to_first_file(self, first_file, row_count):
        """在第一个文件的第一列添加店铺列"""
        if not self.ws or self.has_shop_column:
            return
            
        # 在第一列插入新列
        self.ws.insert_cols(1)
        
        # 设置表头
        shop_header_cell = self.ws.cell(row=1, column=1)
        shop_header_cell.value = "店铺"
        
        # 设置表头格式（复制第二列的格式）
        if self.ws.max_column >= 2:
            ref_cell = self.ws.cell(row=1, column=2)
            FormatHandler.copy_cell_format(ref_cell, shop_header_cell)
        
        # 填充第一个文件的数据来源（文件名，不含扩展名）
        shop_name = os.path.splitext(first_file)[0]
        for row in range(2, row_count + 2):  # 从第二行到数据结束行
            if row > self.ws.max_row:
                break  # 防止超出表格范围
            cell = self.ws.cell(row=row, column=1)
            cell.value = shop_name
            
            # 复制格式（从同行列的原第一列，现在是第二列）
            if self.ws.max_column >= 2:
                ref_cell = self.ws.cell(row=row, column=2)
                FormatHandler.copy_cell_format(ref_cell, cell)
        
        # 更新表头信息，将店铺列包含在内
        self.header_info.insert(0, (1, "店铺", "店铺", False))
        self.header_map["店铺"] = 1
        
        # 调整其他列的索引（+1因为插入了新列）
        for i in range(1, len(self.header_info)):
            col_idx, orig_header, normalized, is_amount_col = self.header_info[i]
            self.header_info[i] = (col_idx + 1, orig_header, normalized, is_amount_col)
            if normalized in self.header_map:
                self.header_map[normalized] = col_idx + 1
        
        self.has_shop_column = True
        self.log(f"已添加店铺列，用于标识数据来源文件")

    def _analyze_first_file(self, first_file):
        """分析第一个文件，获取完整表头信息"""
        first_file_path = os.path.join(self.folder_path, first_file)
        self.log(f"以文件 '{first_file}' 为基础获取表头信息")
        
        # 加载第一个文件
        try:
            self.wb = load_workbook(first_file_path)
        except Exception as e:
            raise IOError(f"无法加载文件 {first_file}: {str(e)}")
            
        self.ws = self.wb.active
        
        # 分析表头（获取所有列的完整信息）
        col_idx = 1
        self.header_info = []  # 重置表头信息
        self.header_map = {}   # 重置表头映射
        while True:
            original_header = self.ws.cell(row=1, column=col_idx).value
            normalized = Utils.normalize_header(original_header)
            is_amount_col = Utils.is_amount_column(str(original_header)) if original_header else False
            
            self.header_info.append((col_idx, original_header, normalized, is_amount_col))
            
            if normalized and normalized not in self.header_map:
                self.header_map[normalized] = col_idx
            
            # 检测表头结束或达到最大列数
            if self._is_header_end(col_idx) or col_idx >= Config.MAX_COLUMNS_TO_CHECK:
                break
                
            col_idx += 1
        
        # 显示检测到的所有表头，确保商品ID、货品ID等关键列被识别
        self.log(f"检测到的表头信息（共 {len(self.header_info)} 列）:")
        for idx, orig, norm, _ in self.header_info:
            self.log(f"  第{idx}列: 原始='{orig}'，标准化='{norm}'")
        
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
            if col_idx + i > Config.MAX_COLUMNS_TO_CHECK:
                return True
            if self.ws.cell(row=1, column=col_idx + i).value is None:
                empty_count += 1
        return empty_count >= Config.EMPTY_COLUMN_THRESHOLD

    def _process_all_files_in_native_order(self):
        """按目录原生顺序处理所有文件并合并数据 - 确保数据完整"""
        all_data = []
        
        for file_idx, file in enumerate(self.excel_files):
            file_path = os.path.join(self.folder_path, file)
            try:
                # 读取文件数据，保留所有列
                df = pd.read_excel(file_path, dtype=str)
                self.log(f"\n处理第{file_idx + 1}个文件: {file} (共{len(df)}行数据)")
                self.log(f"  文件包含列: {', '.join(df.columns.tolist())}")
                
                # 提取店铺名（文件名不含扩展名）
                shop_name = os.path.splitext(file)[0]
                
                # 在第一列插入店铺列
                df.insert(0, '店铺', shop_name)
                
                # 创建对齐后的DataFrame，确保包含所有表头列
                aligned_df = pd.DataFrame(columns=[h[2] for h in self.header_info])
                
                # 逐列映射，确保不丢失任何数据
                for norm_header in aligned_df.columns:
                    # 尝试找到最匹配的列
                    matched = False
                    for df_col in df.columns:
                        if Utils.normalize_header(df_col) == norm_header:
                            aligned_df[norm_header] = df[df_col]
                            matched = True
                            break
                    
                    # 如果未找到匹配列，保持为空但保留列
                    if not matched and norm_header != '店铺':
                        self.log(f"  警告: 文件中未找到与 '{norm_header}' 匹配的列，将保留空值")
                        aligned_df[norm_header] = ""
                
                all_data.append(aligned_df)
                self.log(f"  处理完成，已映射所有列")
                
            except Exception as e:
                self.log(f"警告: 处理文件{file}时出错，已跳过 - {str(e)}")
        
        if not all_data:
            raise ValueError("没有可处理的有效文件")
        
        # 合并所有数据（严格保持文件的原生顺序）
        merged_df = pd.concat(all_data, ignore_index=True)
        self.log(f"\n数据合并完成，共 {len(merged_df)} 行数据，{len(merged_df.columns)} 列")
        return merged_df

    def _write_merged_data(self, merged_df, first_row_count, start_row, format_ref_row):
        """将合并后的数据写入Excel - 确保所有列数据正确写入"""
        # 清除原有多余数据（保留第一个文件的数据）
        if self.ws.max_row >= start_row:
            try:
                self.ws.delete_rows(start_row, self.ws.max_row - start_row + 1)
                self.log("已清除基础文件后的原有数据")
            except Exception as e:
                self.log(f"警告: 清除旧数据时出错 - {str(e)}")
        
        # 写入数据（从第一个文件之后开始）
        total_rows = len(merged_df)
        batch_size = Config.WRITE_BATCH_SIZE
        
        for batch_start in range(first_row_count, total_rows, batch_size):
            batch_end = min(batch_start + batch_size, total_rows)
            self.log(f"正在写入数据: {batch_end}/{total_rows} 行")
            
            for row_idx in range(batch_start, batch_end):
                data_row = merged_df.iloc[row_idx]
                current_row = start_row + (row_idx - first_row_count)
                
                for col_info in self.header_info:
                    self._write_cell(data_row, col_info, current_row, format_ref_row)

    def _write_cell(self, data_row, col_info, current_row, format_ref_row):
        """写入单个单元格数据 - 确保所有值正确保留"""
        col_idx, orig_header, norm_header, is_amount_col = col_info
        
        # 获取单元格值（确保不丢失数据）
        try:
            value = data_row[norm_header] if norm_header in data_row else ""
            if pd.isna(value):
                value = ""
        except:
            value = ""
        
        # 获取参考单元格和目标单元格
        ref_cell = self.ws.cell(row=format_ref_row, column=col_idx)
        target_cell = self.ws.cell(row=current_row, column=col_idx)
        
        # 处理金额列
        if is_amount_col:
            value = self._process_amount_value(value, target_cell, ref_cell)
        # 处理长数字（如ID），确保完整显示
        elif str(value).isdigit() and len(str(value)) > Config.LONG_NUMBER_THRESHOLD:
            target_cell.number_format = '@'  # 文本格式
            value = str(value)
        else:
            target_cell.number_format = ref_cell.number_format
        
        target_cell.value = value
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
