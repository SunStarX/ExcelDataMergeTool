"""图形界面窗口实现 - 完整显示所有按钮"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from .processor import ExcelProcessor

class FinancialDataMergerGUI:
    """带GUI界面的财务数据汇总工具"""
    def __init__(self, root):
        # 配置主窗口
        self.root = root
        self.root.title("财务数据汇总工具")
        self.root.geometry("800x500")  # 确保所有元素可见
        self.root.resizable(False, False)
        
        # 软件信息
        self.app_info = {
            "name": "财务数据汇总工具",
            "version": "1.1.0",
            "author": "sunstar",
            "date": "2025年8月10日"
        }
        
        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        self.style.configure("TLabelframe", font=("SimHei", 10, "bold"))
        
        # 选择的文件夹路径 - 默认为当前工作目录
        current_dir = os.getcwd()
        self.folder_path = tk.StringVar(value=current_dir)
        self.output_path = tk.StringVar(value=current_dir)
        
        # 创建界面
        self._create_widgets()
        
        # 状态变量
        self.processing = False

    def _create_widgets(self):
        """创建所有GUI组件，确保按钮完整显示"""
        # 顶部标题区域
        header_frame = ttk.Frame(self.root, padding="10")
        header_frame.pack(fill=tk.X)
        
        # 标题文字
        ttk.Label(
            header_frame, 
            text="财务数据汇总工具", 
            font=("SimHei", 12, "bold")
        ).pack(side=tk.LEFT)
        
        # 关于按钮
        ttk.Button(
            header_frame, 
            text="关于", 
            command=self._show_about
        ).pack(side=tk.RIGHT)
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 文件夹选择区域
        folder_frame = ttk.LabelFrame(main_frame, text="Excel文件所在文件夹", padding="10")
        folder_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(
            folder_frame, 
            textvariable=self.folder_path, 
            width=50
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            folder_frame, 
            text="浏览...", 
            command=self._browse_folder
        ).pack(side=tk.RIGHT)
        
        # 输出文件夹选择区域
        output_frame = ttk.LabelFrame(main_frame, text="汇总结果保存位置", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(
            output_frame, 
            textvariable=self.output_path, 
            width=50
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(
            output_frame, 
            text="浏览...", 
            command=self._browse_output
        ).pack(side=tk.RIGHT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 日志文本框
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 按钮区域 - 确保按钮完整显示
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(fill=tk.X)
        
        # 明确创建两个按钮并确保它们被正确放置
        self.merge_btn = ttk.Button(
            button_frame, 
            text="开始汇总", 
            command=self._start_merge,
            width=15
        )
        self.merge_btn.pack(side=tk.RIGHT, padx=10)
        
        self.clear_btn = ttk.Button(
            button_frame, 
            text="清空日志", 
            command=self._clear_log,
            width=10
        )
        self.clear_btn.pack(side=tk.RIGHT)

    def _show_about(self):
        """显示关于对话框"""
        about_info = (
            f"{self.app_info['name']}\n\n"
            f"版本: {self.app_info['version']}\n"
            f"作者: {self.app_info['author']}\n"
            f"日期: {self.app_info['date']}"
        )
        messagebox.showinfo("关于", about_info)

    def _browse_folder(self):
        """浏览选择文件夹"""
        folder = filedialog.askdirectory(title="选择Excel文件所在文件夹")
        if folder:
            self.folder_path.set(folder)
    
    def _browse_output(self):
        """浏览选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择汇总结果保存位置")
        if folder:
            self.output_path.set(folder)
    
    def _log(self, message, is_error=False):
        """在日志区域显示消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def _clear_log(self):
        """清空日志"""
        if messagebox.askyesno("确认", "确定要清空日志吗？"):
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
    
    def _start_merge(self):
        """开始汇总过程"""
        if self.processing:
            messagebox.showinfo("提示", "正在处理中，请稍候...")
            return
            
        folder_path = self.folder_path.get()
        output_path = self.output_path.get()
        
        # 验证路径
        if not folder_path:
            messagebox.showerror("错误", "请选择Excel文件所在文件夹")
            return
            
        if not output_path:
            messagebox.showerror("错误", "请选择结果保存位置")
            return
            
        if not os.path.exists(folder_path):
            messagebox.showerror("错误", f"文件夹不存在: {folder_path}")
            return
            
        if not os.path.exists(output_path):
            messagebox.showerror("错误", f"输出文件夹不存在: {output_path}")
            return
        
        # 开始处理
        self.processing = True
        self._log("开始汇总Excel文件...")
        self.merge_btn.config(state=tk.DISABLED)
        self.clear_btn.config(state=tk.DISABLED)
        
        try:
            processor = ExcelProcessor(folder_path, output_path, self._log)
            result_file = processor.merge()
            
            if result_file:
                self._log(f"汇总成功！结果已保存至:\n{result_file}")
                if messagebox.askyesno("成功", "汇总成功！是否打开结果文件所在文件夹？"):
                    os.startfile(os.path.dirname(result_file))
        except Exception as e:
            self._log(f"汇总失败: {str(e)}")
            messagebox.showerror("错误", f"汇总失败: {str(e)}")
        finally:
            # 恢复按钮状态
            self.processing = False
            self.merge_btn.config(state=tk.NORMAL)
            self.clear_btn.config(state=tk.NORMAL)
