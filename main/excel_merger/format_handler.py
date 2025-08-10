"""单元格格式处理工具 - 避免StyleProxy哈希操作"""
from openpyxl.styles import Font, Alignment, Border, PatternFill

class FormatHandler:
    """Excel单元格格式处理类 - 修复StyleProxy哈希错误"""

    @staticmethod
    def copy_cell_format(source_cell, target_cell, force_right=False):
        """
        复制单元格格式从源单元格到目标单元格
        避免直接使用StyleProxy对象进行哈希操作
        """
        if not source_cell or not target_cell:
            return

        # 复制字体格式 - 提取具体属性而非使用整个StyleProxy
        target_cell.font = Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            vertAlign=source_cell.font.vertAlign,
            underline=source_cell.font.underline,
            strike=source_cell.font.strike,
            color=source_cell.font.color
        )

        # 复制对齐方式
        horizontal = "right" if force_right else source_cell.alignment.horizontal
        target_cell.alignment = Alignment(
            horizontal=horizontal,
            vertical=source_cell.alignment.vertical,
            text_rotation=source_cell.alignment.text_rotation,
            wrap_text=source_cell.alignment.wrap_text,
            shrink_to_fit=source_cell.alignment.shrink_to_fit,
            indent=source_cell.alignment.indent
        )

        # 复制边框
        target_cell.border = Border(
            left=source_cell.border.left,
            right=source_cell.border.right,
            top=source_cell.border.top,
            bottom=source_cell.border.bottom,
            diagonal=source_cell.border.diagonal,
            diagonal_direction=source_cell.border.diagonal_direction,
            outline=source_cell.border.outline,
            vertical=source_cell.border.vertical,
            horizontal=source_cell.border.horizontal
        )

        # 复制填充颜色
        target_cell.fill = PatternFill(
            patternType=source_cell.fill.patternType,
            fgColor=source_cell.fill.fgColor,
            bgColor=source_cell.fill.bgColor
        )

        # 复制数字格式
        target_cell.number_format = source_cell.number_format
