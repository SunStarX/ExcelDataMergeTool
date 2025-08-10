"""项目运行入口"""
import tkinter as tk
from tkinter import messagebox
from excel_merger.gui_window import FinancialDataMergerGUI

def main():
    """启动应用程序"""
    try:
        root = tk.Tk()
        # 尝试设置窗口图标（可选）
        try:
            root.iconbitmap(default="")  # 可替换为实际图标路径
        except:
            pass
        app = FinancialDataMergerGUI(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("启动错误", f"程序启动失败: {str(e)}")

if __name__ == "__main__":
    main()
