"""单元格格式处理工具"""

from openpyxl.styles import Alignment, Font, PatternFill, Border

class FormatHandler:
    """单元格格式处理工具"""
    @staticmethod
    def copy_cell_format(source_cell, target_cell, force_right: bool = False) -> None:
        """复制单元格格式，可强制设置右对齐"""
        # 处理对齐方式
        horizontal = source_cell.alignment.horizontal
        if force_right:
            horizontal = "right"
            
        target_cell.alignment = Alignment(
            horizontal=horizontal,
            vertical=source_cell.alignment.vertical,
            textRotation=source_cell.alignment.textRotation,
            wrapText=source_cell.alignment.wrapText,
            shrinkToFit=source_cell.alignment.shrinkToFit,
            indent=source_cell.alignment.indent
        )
        
        # 复制其他格式
        target_cell.font = Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            color=source_cell.font.color,
            underline=source_cell.font.underline
        )
        
        target_cell.fill = PatternFill(
            patternType=source_cell.fill.patternType,
            fgColor=source_cell.fill.fgColor,
            bgColor=source_cell.fill.bgColor
        )
        
        target_cell.border = Border(
            left=source_cell.border.left,
            right=source_cell.border.right,
            top=source_cell.border.top,
            bottom=source_cell.border.bottom
        )
        
        target_cell.number_format = source_cell.number_format
