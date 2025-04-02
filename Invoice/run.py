import os
import sys
import traceback
from tkinter import messagebox
import tkinter as tk

def show_error(error_msg):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("错误", error_msg)
    root.destroy()

def main():
    try:
        # 将当前目录添加到 Python 路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)

        import invoice_gui
        
        # 创建日志目录
        if not os.path.exists('logs'):
            os.makedirs('logs')
            
        # 启动主程序
        root = tk.Tk()
        app = invoice_gui.InvoiceProcessorGUI(root)
        root.mainloop()
        
    except Exception as e:
        error_msg = f"程序启动失败!\n\n错误信息:\n{str(e)}\n\n详细信息:\n{traceback.format_exc()}"
        show_error(error_msg)
        
        # 保存错误日志
        try:
            with open('logs/error.log', 'w', encoding='utf-8') as f:
                f.write(error_msg)
        except:
            pass
            
        # 等待用户输入后再退出
        input("\n按回车键退出...")

if __name__ == "__main__":
    main() 