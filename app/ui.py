import tkinter as tk
from tkinter import ttk, messagebox
from logic import export_report


def export_action(template_var, report_type_var, export_format_var):
    template = template_var.get()
    report_type = report_type_var.get()
    export_format = export_format_var.get()

    if template and report_type and export_format:
        try:
            export_report(template, report_type, export_format)
            messagebox.showinfo("Success", "报告导出成功")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Warning", "请确保所有选项都已选择")


def center_window(root, width, height):
    # 获取屏幕 宽、高
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()

    # 计算窗口居中的位置坐标
    alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)

    # 设置窗口大小和位置
    root.geometry(alignstr)
    root.resizable(width=False, height=False)  # 禁止用户改变窗口大小


def main_ui():
    root = tk.Tk()
    center_window(root, 500, 400)  # 传入窗口对象和窗口的宽高
    root.title("Easy Audit")
    # root.geometry("500x400")
    root.configure(bg="#f5f5f5")

    # 标签框提示文本样式
    label_style = {"font": ("Arial", 10), "bg": "#f5f5f5"}

    # 下拉框变量
    template_var = tk.StringVar()
    report_type_var = tk.StringVar()
    export_format_var = tk.StringVar()

    # 下拉选择框及其标签
    tk.Label(root, text="请先选择报告模板", **label_style).pack(pady=(20, 5))
    template_options = ["高新", "普通"]
    template_menu = ttk.Combobox(root, textvariable=template_var, values=template_options, state="readonly")
    template_menu.pack(pady=(0, 20))

    tk.Label(root, text="请先选择报告类型", **label_style).pack(pady=(0, 5))
    report_type_options = ["年报", "费用专项", "收入专项"]
    report_type_menu = ttk.Combobox(root, textvariable=report_type_var, values=report_type_options, state="readonly")
    report_type_menu.pack(pady=(0, 20))

    tk.Label(root, text="请先选择导出格式", **label_style).pack(pady=(0, 5))
    export_format_options = ["PDF", "Word"]
    export_format_menu = ttk.Combobox(root, textvariable=export_format_var, values=export_format_options,
                                      state="readonly")
    export_format_menu.pack(pady=(0, 20))

    # 导出按钮
    export_button = tk.Button(root, text="导出",
                              command=lambda: export_action(template_var, report_type_var, export_format_var),
                              font=("Arial", 13), bg="#4CAF50", fg="white", bd=2, relief="flat", width=17, height=1)
    export_button.pack(pady=30)

    # 按钮样式美化
    def on_enter(e):
        export_button.config(bg="#45a049")

    def on_leave(e):
        export_button.config(bg="#4CAF50")

    export_button.bind("<Enter>", on_enter)
    export_button.bind("<Leave>", on_leave)

    # 运行主循环
    root.mainloop()


if __name__ == "__main__":
    main_ui()
