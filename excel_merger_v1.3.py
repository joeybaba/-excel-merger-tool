#!/usr/bin/env python3
import sys
import time
import psutil
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import xlwings as xw
import chardet
import threading
from queue import Queue

class ExcelMerger:
    def __init__(self):
        self.main_file = None
        self.sub_files = {}
        self.main_data = {}
        self.sub_data = {}
        self.progress_queue = Queue()
        self.setup_gui()

    def detect_encoding(self, file_path):
        """检测文件编码"""
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            return result['encoding']

    def load_excel_file(self, file_path):
        """通用的Excel文件加载函数"""
        start_time = time.time()
        try:
            # 首先尝试使用pandas直接读取
            if file_path.lower().endswith('.ods'):
                df = pd.read_excel(file_path, engine='odf')
            else:
                df = pd.read_excel(file_path)
            return df
        except Exception as e:
            try:
                # 如果失败，尝试使用xlwings读取
                app = xw.App(visible=False)
                wb = app.books.open(file_path)
                df = wb.sheets[0].used_range.options(pd.DataFrame, index=False).value
                wb.close()
                app.quit()
                return df
            except Exception as e2:
                # 尝试使用其他引擎读取
                try:
                    engines = ['openpyxl', 'xlrd', 'pyxlsb']
                    for engine in engines:
                        try:
                            df = pd.read_excel(file_path, engine=engine)
                            return df
                        except Exception:
                            continue
                    raise Exception("所有读取方法都失败")
                except Exception as e3:
                    raise Exception(f"无法读取文件，请确保文件格式正确。\n错误详情：\n{str(e)}\n{str(e2)}\n{str(e3)}")
            finally:
                process = psutil.Process()

    def on_drop_main(self, event):
        """处理主表文件的拖放事件"""
        file_path = event.data
        if file_path.startswith('{'): # Windows 路径处理
            import json
            file_path = json.loads(file_path)['text']
        if file_path.startswith('file://'): # macOS 路径处理
            file_path = file_path.replace('file://', '')
        if os.path.isfile(file_path) and file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.et', '.ett')):
            self.load_main_file(file_path)  # 直接传递文件路径
        else:
            messagebox.showerror("错误", "请拖放有效的Excel文件！")

    def on_drop_sub(self, event, sheet_name):
        """处理副表文件的拖放事件"""
        file_path = event.data
        if file_path.startswith('{'): # Windows 路径处理
            import json
            file_path = json.loads(file_path)['text']
        if file_path.startswith('file://'): # macOS 路径处理
            file_path = file_path.replace('file://', '')
        if os.path.isfile(file_path):
            try:
                self.update_status(f"正在加载{sheet_name}的副表文件...")
                self.progress_var.set(0)

                # 初始化或重置该工作表的副表数据
                if sheet_name not in self.sub_data:
                    self.sub_data[sheet_name] = pd.DataFrame()
                    self.sub_files[sheet_name] = []

                # 加载拖放的文件
                if file_path.lower().endswith('.csv'):
                    try:
                        df = pd.read_csv(file_path)
                    except UnicodeDecodeError:
                        encoding = self.detect_encoding(file_path)
                        df = pd.read_csv(file_path, encoding=encoding)
                else:
                    df = self.load_excel_file(file_path)

                # 将当前文件的数据添加到该工作表的副表数据中
                self.sub_data[sheet_name] = pd.concat([self.sub_data[sheet_name], df], ignore_index=True)
                self.sub_files[sheet_name].append(file_path)

                self.progress_var.set(100)
                total_rows = len(self.sub_data[sheet_name])
                loaded_files = "\n".join([os.path.basename(f) for f in self.sub_files[sheet_name]])
                self.update_status(f"{sheet_name}副表文件加载成功\n已加载的文件：\n{loaded_files}\n总行数：{total_rows}")

            except Exception as e:
                messagebox.showerror("错误", f"加载文件 {os.path.basename(file_path)} 时出错：\n{str(e)}")
        else:
            messagebox.showerror("错误", "请拖放有效的Excel或CSV文件！")

    def update_status(self, message):
        """更新状态信息"""
        print(message)  # 控制台输出
        self.status_label.config(text=message)  # GUI更新
        self.root.update()

    def clear_all_files(self):
        """清理所有已加载的文件数据"""
        self.main_file = None
        self.sub_files = {}
        self.main_data = {}
        self.sub_data = {}
        print("已清理所有已加载的文件数据")
        self.update_status("已清理所有文件，请重新选择文件")

    def setup_gui(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Excel文件合并工具")
        self.root.geometry("600x800")

        # 创建主框架，增加内边距
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        # 创建文件选择区域，增加内边距和间距
        file_frame = ttk.LabelFrame(main_frame, text="主表文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=10)

        # 主表文件选择区域，优化布局
        main_file_frame = ttk.Frame(file_frame)
        main_file_frame.pack(fill=tk.X, pady=8)
        select_main_btn = ttk.Button(main_file_frame, text="选择主表文件", command=lambda: self.load_main_file(None))
        select_main_btn.pack(side=tk.LEFT, padx=10)
        
        # 主表拖拽区域，增加高度和内边距
        main_drop_frame = ttk.LabelFrame(main_file_frame, text="拖拽区域", width=200, height=50)
        main_drop_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        main_drop_frame.drop_target_register("DND_Files")
        main_drop_frame.dnd_bind('<<Drop>>', self.on_drop_main)
        ttk.Label(main_drop_frame, text="拖拽文件到这里").pack(pady=8)

        # 副表选择区域，增加内边距和间距
        sub_files_frame = ttk.LabelFrame(main_frame, text="副表文件选择", padding=10)
        sub_files_frame.pack(fill=tk.X, pady=10)

        # 创建副表选择的函数
        def create_sub_file_frame(parent, text, var, sheet_name):
            frame = ttk.Frame(parent)
            frame.pack(fill=tk.X, pady=8)
            ttk.Checkbutton(frame, text=text, variable=var).pack(side=tk.LEFT, padx=10)
            ttk.Button(frame, text="选择副表", command=lambda: self.load_sub_file(sheet_name)).pack(side=tk.LEFT, padx=10)
            
            drop_frame = ttk.LabelFrame(frame, text="拖拽区域", width=200, height=50)
            drop_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
            drop_frame.drop_target_register("DND_Files")
            drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, sheet_name))
            ttk.Label(drop_frame, text="拖拽文件到这里").pack(pady=8)

        # 使用函数创建各个副表框架
        create_sub_file_frame(sub_files_frame, "全站营销", self.merge_marketing, "全站营销")
        create_sub_file_frame(sub_files_frame, "站内数据源", self.merge_internal, "站内数据源")
        create_sub_file_frame(sub_files_frame, "站外数据源", self.merge_external, "站外数据源")
        create_sub_file_frame(sub_files_frame, "店铺成交数据源", self.merge_shop, "店铺成交数据源")

        # 合并按钮，增加样式和间距
        merge_button = ttk.Button(main_frame, text="合并文件", command=self.merge_files)
        merge_button.pack(pady=15)

        # 进度条，增加内边距
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=10, padx=5)

        # 状态标签，增加宽度和内边距
        self.status_label = ttk.Label(main_frame, text="请选择文件", wraplength=550, justify=tk.LEFT)
        self.status_label.pack(pady=10, fill=tk.X)

        self.root.mainloop()

    def load_main_file(self, file_path=None):
        """加载主表文件"""
        # 如果是拖拽的文件路径，直接使用；如果是通过按钮点击，则显示文件选择对话框
        if file_path is None:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.et *.ett")])
            if not file_path:  # 如果用户取消选择，直接返回
                return
        elif isinstance(file_path, tuple):
            file_path = file_path[0]
        
        try:
            self.update_status("正在加载主表文件...")
            self.progress_var.set(0)
            
            # 使用pandas读取文件，检查工作表
            try:
                excel_file = pd.ExcelFile(file_path)
                required_sheets = []
                if self.merge_marketing.get():
                    required_sheets.append("全站营销")
                if self.merge_internal.get():
                    required_sheets.append("站内数据源")
                if self.merge_external.get():
                    required_sheets.append("站外数据源")
                if self.merge_shop.get():
                    required_sheets.append("店铺成交数据源")
                
                missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
                if missing_sheets:
                    messagebox.showerror("错误", f"主表文件中未找到以下工作表：\n{', '.join(missing_sheets)}\n请确保文件包含正确的工作表。")
                    return

                # 读取选中的工作表
                progress_step = 50 / len(required_sheets)
                for sheet_name in required_sheets:
                    self.main_data[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
                    self.progress_var.set(self.progress_var.get() + progress_step)

            except Exception as e:
                # 如果pandas读取失败，尝试使用xlwings
                try:
                    app = xw.App(visible=False)
                    wb = app.books.open(file_path)
                    sheet_names = [sheet.name for sheet in wb.sheets]
                    
                    required_sheets = []
                    if self.merge_marketing.get():
                        required_sheets.append("全站营销")
                    if self.merge_internal.get():
                        required_sheets.append("站内数据源")
                    if self.merge_external.get():
                        required_sheets.append("站外数据源")
                    if self.merge_shop.get():
                        required_sheets.append("店铺成交数据源")
                    
                    missing_sheets = [sheet for sheet in required_sheets if sheet not in sheet_names]
                    if missing_sheets:
                        wb.close()
                        app.quit()
                        messagebox.showerror("错误", f"主表文件中未找到以下工作表：\n{', '.join(missing_sheets)}\n请确保文件包含正确的工作表。")
                        return

                    # 读取选中的工作表
                    progress_step = 50 / len(required_sheets)
                    for sheet_name in required_sheets:
                        self.main_data[sheet_name] = wb.sheets[sheet_name].used_range.options(pd.DataFrame, index=False).value
                        self.progress_var.set(self.progress_var.get() + progress_step)

                    wb.close()
                    app.quit()

                except Exception as e2:
                    messagebox.showerror("错误", f"无法读取主表文件，请确保文件格式正确。\n错误详情：\n{str(e)}\n{str(e2)}")
                    return

            self.main_file = file_path
            self.progress_var.set(100)
            sheet_info = "\n".join([f"{sheet}：{len(data)}行" for sheet, data in self.main_data.items()])
            self.update_status(f"主表文件加载成功\n文件路径：{file_path}\n{sheet_info}")
        except Exception as e:
            messagebox.showerror("错误", f"加载主表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏")

    def load_sub_file(self, sheet_name):
        """加载指定工作表的副表文件"""
        # 只有当通过按钮点击时才显示文件选择对话框
        file_paths = filedialog.askopenfilenames(filetypes=[
            ("All supported files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett *.csv"), 
            ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett"), 
            ("CSV files", "*.csv"), 
            ("All files", "*.*")])
        
        if not file_paths:  # 如果用户取消选择，直接返回
            return

        try:
            self.update_status(f"正在加载{sheet_name}的副表文件...")
            self.progress_var.set(0)
            progress_step = 100 / len(file_paths)
            current_progress = 0

            # 初始化或重置该工作表的副表数据
            if sheet_name not in self.sub_data:
                self.sub_data[sheet_name] = pd.DataFrame()
                self.sub_files[sheet_name] = []

            # 循环处理每个选中的文件
            for file_path in file_paths:
                try:
                    if file_path.lower().endswith('.csv'):
                        try:
                            # 首先尝试使用系统默认编码读取
                            df = pd.read_csv(file_path)
                        except UnicodeDecodeError:
                            # 如果失败，尝试检测文件编码
                            encoding = self.detect_encoding(file_path)
                            df = pd.read_csv(file_path, encoding=encoding)
                    else:
                        # 尝试使用通用Excel加载函数
                        df = self.load_excel_file(file_path)

                    # 将当前文件的数据添加到该工作表的副表数据中
                    self.sub_data[sheet_name] = pd.concat([self.sub_data[sheet_name], df], ignore_index=True)
                    self.sub_files[sheet_name].append(file_path)

                    current_progress += progress_step
                    self.progress_var.set(current_progress)
                    self.update_status(f"已加载文件：{os.path.basename(file_path)}")

                except Exception as e:
                    messagebox.showerror("错误", f"加载文件 {os.path.basename(file_path)} 时出错：\n{str(e)}")
                    continue

            self.progress_var.set(100)
            total_rows = len(self.sub_data[sheet_name])
            loaded_files = "\n".join([os.path.basename(f) for f in self.sub_files[sheet_name]])
            self.update_status(f"{sheet_name}副表文件加载成功\n已加载的文件：\n{loaded_files}\n总行数：{total_rows}")

        except Exception as e:
            messagebox.showerror("错误", f"加载副表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏\n4. CSV文件编码格式正确")

    def setup_gui(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Excel文件合并工具")
        self.root.geometry("400x600")

        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 创建文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding=5)
        file_frame.pack(fill=tk.X, pady=5)

        # 主表文件选择区域
        main_file_frame = ttk.Frame(file_frame)
        main_file_frame.pack(fill=tk.X, pady=5)
        button_frame = ttk.Frame(main_file_frame)
        button_frame.pack(side=tk.LEFT, padx=(5, 15))
        ttk.Button(button_frame, text="选择主表文件", command=lambda: self.load_main_file(None)).pack()
        
        # 主表拖拽区域
        main_drop_frame = ttk.LabelFrame(main_file_frame, width=200, height=35)
        main_drop_frame.pack(side=tk.RIGHT, padx=(0, 5), fill=tk.X, expand=True)
        main_drop_frame.drop_target_register("DND_Files")
        main_drop_frame.dnd_bind('<<Drop>>', self.on_drop_main)
        ttk.Label(main_drop_frame, text="拖拽文件到这里").pack(pady=3)

        # 副表选择区域
        sub_files_frame = ttk.LabelFrame(main_frame, text="副表文件选择", padding=5)
        sub_files_frame.pack(fill=tk.X, pady=5)

        # 全站营销副表
        marketing_frame = ttk.Frame(sub_files_frame)
        marketing_frame.pack(fill=tk.X, pady=2)
        left_frame = ttk.Frame(marketing_frame)
        left_frame.pack(side=tk.LEFT, padx=(5, 15))
        self.merge_marketing = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_frame, text="全站营销", variable=self.merge_marketing).pack(side=tk.LEFT)
        ttk.Button(left_frame, text="选择副表", 
                   command=lambda: self.load_sub_file("全站营销")).pack(side=tk.LEFT, padx=5)
        
        # 全站营销拖拽区域
        marketing_drop_frame = ttk.LabelFrame(marketing_frame, width=200, height=35)
        marketing_drop_frame.pack(side=tk.RIGHT, padx=(0, 5), fill=tk.X, expand=True)
        marketing_drop_frame.drop_target_register("DND_Files")
        marketing_drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, "全站营销"))
        ttk.Label(marketing_drop_frame, text="拖拽文件到这里").pack(pady=3)

        # 站内数据源副表
        internal_frame = ttk.Frame(sub_files_frame)
        internal_frame.pack(fill=tk.X, pady=2)
        left_frame = ttk.Frame(internal_frame)
        left_frame.pack(side=tk.LEFT, padx=(5, 15))
        self.merge_internal = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_frame, text="站内数据源", variable=self.merge_internal).pack(side=tk.LEFT)
        ttk.Button(left_frame, text="选择副表", 
                   command=lambda: self.load_sub_file("站内数据源")).pack(side=tk.LEFT, padx=5)
        
        # 站内数据源拖拽区域
        internal_drop_frame = ttk.LabelFrame(internal_frame, width=200, height=35)
        internal_drop_frame.pack(side=tk.RIGHT, padx=(0, 5), fill=tk.X, expand=True)
        internal_drop_frame.drop_target_register("DND_Files")
        internal_drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, "站内数据源"))
        ttk.Label(internal_drop_frame, text="拖拽文件到这里").pack(pady=3)

        # 站外数据源副表
        external_frame = ttk.Frame(sub_files_frame)
        external_frame.pack(fill=tk.X, pady=2)
        left_frame = ttk.Frame(external_frame)
        left_frame.pack(side=tk.LEFT, padx=(5, 15))
        self.merge_external = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_frame, text="站外数据源", variable=self.merge_external).pack(side=tk.LEFT)
        ttk.Button(left_frame, text="选择副表", 
                   command=lambda: self.load_sub_file("站外数据源")).pack(side=tk.LEFT, padx=5)
        
        # 站外数据源拖拽区域
        external_drop_frame = ttk.LabelFrame(external_frame, width=200, height=35)
        external_drop_frame.pack(side=tk.RIGHT, padx=(0, 5), fill=tk.X, expand=True)
        external_drop_frame.drop_target_register("DND_Files")
        external_drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, "站外数据源"))
        ttk.Label(external_drop_frame, text="拖拽文件到这里").pack(pady=3)

        # 店铺成交数据源副表
        shop_frame = ttk.Frame(sub_files_frame)
        shop_frame.pack(fill=tk.X, pady=2)
        left_frame = ttk.Frame(shop_frame)
        left_frame.pack(side=tk.LEFT, padx=(5, 15))
        self.merge_shop = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_frame, text="店铺成交数据源", variable=self.merge_shop).pack(side=tk.LEFT)
        ttk.Button(left_frame, text="选择副表", 
                   command=lambda: self.load_sub_file("店铺成交数据源")).pack(side=tk.LEFT, padx=5)
        
        # 店铺成交数据源拖拽区域
        shop_drop_frame = ttk.LabelFrame(shop_frame, width=200, height=35)
        shop_drop_frame.pack(side=tk.RIGHT, padx=(0, 5), fill=tk.X, expand=True)
        shop_drop_frame.drop_target_register("DND_Files")
        shop_drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, "店铺成交数据源"))
        ttk.Label(shop_drop_frame, text="拖拽文件到这里").pack(pady=3)

        # 合并按钮
        ttk.Button(main_frame, text="合并文件", command=self.merge_files).pack(pady=10)

        # 清理按钮
        ttk.Button(main_frame, text="清理所有文件", command=self.clear_all_files).pack(pady=5)

        # 进度条
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=5)

        # 状态标签
        self.status_label = ttk.Label(main_frame, text="请选择文件", wraplength=350)
        self.status_label.pack(pady=5)

        self.root.mainloop()

    def load_main_file(self, file_path=None):
        """加载主表文件"""
        if file_path is None:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.et *.ett")])
            if not file_path:  # 如果用户取消选择，直接返回
                return
        elif isinstance(file_path, tuple):
            file_path = file_path[0]
        
        try:
            self.update_status("正在加载主表文件...")
            self.progress_var.set(0)
            
            # 使用pandas读取文件，检查工作表
            try:
                excel_file = pd.ExcelFile(file_path)
                required_sheets = []
                if self.merge_marketing.get():
                    required_sheets.append("全站营销")
                if self.merge_internal.get():
                    required_sheets.append("站内数据源")
                if self.merge_external.get():
                    required_sheets.append("站外数据源")
                if self.merge_shop.get():
                    required_sheets.append("店铺成交数据源")
                
                missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
                if missing_sheets:
                    messagebox.showerror("错误", f"主表文件中未找到以下工作表：\n{', '.join(missing_sheets)}\n请确保文件包含正确的工作表。")
                    return

                # 读取选中的工作表
                progress_step = 50 / len(required_sheets)
                for sheet_name in required_sheets:
                    self.main_data[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
                    self.progress_var.set(self.progress_var.get() + progress_step)

            except Exception as e:
                # 如果pandas读取失败，尝试使用xlwings
                try:
                    app = xw.App(visible=False)
                    wb = app.books.open(file_path)
                    sheet_names = [sheet.name for sheet in wb.sheets]
                    
                    required_sheets = []
                    if self.merge_marketing.get():
                        required_sheets.append("全站营销")
                    if self.merge_internal.get():
                        required_sheets.append("站内数据源")
                    if self.merge_external.get():
                        required_sheets.append("站外数据源")
                    if self.merge_shop.get():
                        required_sheets.append("店铺成交数据源")
                    
                    missing_sheets = [sheet for sheet in required_sheets if sheet not in sheet_names]
                    if missing_sheets:
                        wb.close()
                        app.quit()
                        messagebox.showerror("错误", f"主表文件中未找到以下工作表：\n{', '.join(missing_sheets)}\n请确保文件包含正确的工作表。")
                        return

                    # 读取选中的工作表
                    progress_step = 50 / len(required_sheets)
                    for sheet_name in required_sheets:
                        self.main_data[sheet_name] = wb.sheets[sheet_name].used_range.options(pd.DataFrame, index=False).value
                        self.progress_var.set(self.progress_var.get() + progress_step)

                    wb.close()
                    app.quit()

                except Exception as e2:
                    messagebox.showerror("错误", f"无法读取主表文件，请确保文件格式正确。\n错误详情：\n{str(e)}\n{str(e2)}")
                    return

            self.main_file = file_path
            self.progress_var.set(100)
            sheet_info = "\n".join([f"{sheet}：{len(data)}行" for sheet, data in self.main_data.items()])
            self.update_status(f"主表文件加载成功\n文件路径：{file_path}\n{sheet_info}")
        except Exception as e:
            messagebox.showerror("错误", f"加载主表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏")

    def update_status(self, message):
        """更新状态信息"""
        print(message)  # 控制台输出
        self.status_label.config(text=message)  # GUI更新
        self.root.update()

    def clear_all_files(self):
        """清理所有已加载的文件数据"""
        # 重置所有数据
        self.main_file = None
        self.sub_files = {}
        self.main_data = {}
        self.sub_data = {}

        # 重置进度条
        self.progress_var.set(0)

        # 更新状态信息
        self.update_status("所有文件已清理，请重新选择文件")

        print("已清理所有已加载的文件数据")

    def merge_files(self):
        """合并文件"""
        if not self.main_data or not self.sub_data:
            messagebox.showerror("错误", "请先加载主表和副表文件！")
            return

        if not any([self.merge_marketing.get(), self.merge_internal.get(), self.merge_external.get(), self.merge_shop.get()]):
            messagebox.showerror("错误", "请至少选择一个要合并的工作表！")
            return

        try:
            self.update_status("正在处理数据...")
            start_time = time.time()
            self.progress_var.set(0)

            # 使用xlwings打开主表文件以保持公式
            app = xw.App(visible=False)
            wb = app.books.open(self.main_file)
            progress_step = 100 / len(self.sub_data)
            current_progress = 0

            for sheet_name, sub_data in self.sub_data.items():
                # 检查工作表是否被选中
                if sheet_name == "全站营销" and not self.merge_marketing.get():
                    continue
                if sheet_name == "站内数据源" and not self.merge_internal.get():
                    continue
                if sheet_name == "站外数据源" and not self.merge_external.get():
                    continue
                if sheet_name == "店铺成交数据源" and not self.merge_shop.get():
                    continue
                
                if sheet_name not in self.main_data:
                    continue

                main_sheet = wb.sheets[sheet_name]
                main_range = main_sheet.used_range
                # 根据工作表名称决定起始列偏移
                if sheet_name == "站外数据源":
                    start_col_offset = 5  # F列
                elif sheet_name == "店铺成交数据源":
                    start_col_offset = 7  # H列
                else:
                    start_col_offset = 6  # G列
                raw_data = main_range.offset(0, start_col_offset).options(pd.DataFrame, index=False, expand='table').value
                main_data_from_g = pd.DataFrame(raw_data)
                original_columns_count = len(main_data_from_g.columns)

                # 如果是店铺成交数据源工作表，从文件名中提取日期并填充到G列
                if sheet_name == "店铺成交数据源":
                    # 获取主表当前数据的行数和副表的行数
                    main_data_rows = len(main_data_from_g)
                    sub_data_rows = len(sub_data)
                    # 计算G列需要填充的起始行和结束行
                    start_row = main_data_rows + 2  # 从主表数据后的下一行开始
                    
                    # 从每个文件名中提取日期信息并填充对应的行
                    import re
                    g_col_data = []
                    current_row = 0
                    
                    for file_path in self.sub_files[sheet_name]:
                        # 读取当前文件的数据
                        if file_path.lower().endswith('.csv'):
                            try:
                                df = pd.read_csv(file_path)
                            except UnicodeDecodeError:
                                encoding = self.detect_encoding(file_path)
                                df = pd.read_csv(file_path, encoding=encoding)
                        else:
                            df = self.load_excel_file(file_path)
                        
                        # 从文件名中提取日期
                        file_name = os.path.basename(file_path)
                        date_match = re.search(r'_([0-9]{8})_', file_name)
                        if date_match:
                            date_value = int(date_match.group(1))
                        else:
                            logger.warning(f"未能从文件名中提取到日期信息: {file_name}，使用固定值'error'")
                            date_value = "error"
                        
                        # 为当前文件的每一行数据添加对应的日期
                        file_rows = len(df)
                        g_col_data.extend([[date_value] for _ in range(file_rows)])
                        current_row += file_rows
                    
                    # 填充G列数据
                    end_row = start_row + len(g_col_data) - 1
                    main_sheet.range(f"G{start_row}:G{end_row}").value = g_col_data

                # 优化列名处理
                if main_data_from_g.columns.isnull().any():
                    main_data_from_g.columns = [f"Column_{i+1}" if pd.isnull(col) else col 
                                               for i, col in enumerate(main_data_from_g.columns)]

                # 优化空列处理，确保只有当整列（包括表头）都为空时才删除
                empty_columns = main_data_from_g.isna().all() & main_data_from_g.columns.isna()
                if empty_columns.any():
                    # 记录非空列的索引
                    valid_columns = ~empty_columns
                    # 确保至少保留一列数据
                    if valid_columns.any():
                        main_data_from_g = main_data_from_g.loc[:, valid_columns]
                    else:
                        # 如果所有列都是空的，保持原样
                        logger.warning(f"{sheet_name}工作表所有列都是空的，保持原始数据结构")

                # 优化空行处理，只处理连续的空行
                empty_rows = main_data_from_g.isna().all(axis=1)
                if empty_rows.any():
                    # 找到第一个完全空行的索引
                    first_empty_row_idx = empty_rows.idxmax()
                    # 检查是否所有后续行都是空行
                    if empty_rows[first_empty_row_idx:].all():
                        main_data_from_g = main_data_from_g.iloc[:first_empty_row_idx]
                        logger.info(f"从第{first_empty_row_idx}行开始检测到连续空行，已截断数据")

                # 检查数据有效性
                if main_data_from_g.shape[1] == 0:
                    wb.close()
                    app.quit()
                    messagebox.showerror("错误", f"{sheet_name}工作表从G列开始没有有效的数据列，请检查数据格式！")
                    return

                # 检查列数匹配，使用原始列数进行比较
                if original_columns_count != len(sub_data.columns):
                    wb.close()
                    app.quit()
                    messagebox.showerror("错误", f"{sheet_name}工作表列数不匹配: 主表={original_columns_count}, 副表={len(sub_data.columns)}")
                    return

                # 重命名副表列名
                sub_data.columns = main_data_from_g.columns

                # 直接合并数据，不进行重复性检查

                # 合并数据
                merged_data_from_g = pd.concat([main_data_from_g, sub_data], 
                                             axis=0, copy=False).reset_index(drop=True)

                # 找到指定列最后一个非空单元格的位置
                if sheet_name == "站外数据源":
                    start_col = "F"
                elif sheet_name == "店铺成交数据源":
                    start_col = "H"
                else:
                    start_col = "G"
                
                # 获取指定列的所有值
                col_values = main_sheet.range(f"{start_col}1").expand('down').value
                
                # 找到最后一个非空单元格的位置
                last_non_empty_row = 1  # 默认从第一行开始
                for i, value in enumerate(col_values, 1):
                    if value is not None and str(value).strip() != "":
                        last_non_empty_row = i
                
                # 在最后一个非空单元格后追加新数据
                append_start_row = last_non_empty_row + 1

                # 根据工作表名称决定数据写入的起始列
                if sheet_name == "站外数据源":
                    start_col = "F"
                elif sheet_name == "店铺成交数据源":
                    start_col = "H"
                else:
                    start_col = "G"
                main_sheet.range(f"{start_col}{append_start_row}").options(index=False, header=False).value = sub_data.values

                current_progress += progress_step
                self.progress_var.set(current_progress)

            try:
                # 保存文件
                wb.save()
                wb.close()
                app.quit()

                total_time = time.time() - start_time
                self.update_status(f"合并完成！\n数据已保存至原始文件：{self.main_file}\n处理耗时：{total_time:.2f}秒")
                messagebox.showinfo("成功", "数据已成功合并并保存至原始文件")
                self.progress_var.set(100)

            except Exception as save_error:
                logger.error(f"保存原文件失败: {str(save_error)}")
                wb.close()
                app.quit()
                
                # 如果保存失败，创建新文件
                save_path = os.path.join(os.path.dirname(self.main_file), 
                                        f"合并结果_{os.path.basename(self.main_file)}")
                
                # 复制并保存新文件
                import shutil
                shutil.copy2(self.main_file, save_path)
                new_app = xw.App(visible=False)
                new_wb = new_app.books.open(save_path)

                for sheet_name, sub_data in self.sub_data.items():
                    if sheet_name not in self.main_data:
                        continue

                    new_sheet = new_wb.sheets[sheet_name]
                    last_row = new_sheet.used_range.rows.count
                    if last_row > 1:
                        # 根据工作表名称决定起始列
                        if sheet_name == "站外数据源":
                            start_col = "F"
                        elif sheet_name == "店铺成交数据源":
                            start_col = "H"
                        else:
                            start_col = "G"
                        new_sheet.range(f"{start_col}2:XFD{last_row}").clear()

                    # 根据工作表名称决定数据读取的起始列
                    if sheet_name == "站外数据源":
                        start_col = "F"
                    elif sheet_name == "店铺成交数据源":
                        start_col = "H"
                    else:
                        start_col = "G"
                    main_data_from_g = pd.DataFrame(new_sheet.range(f"{start_col}1").expand().value)
                    merged_data_from_g = pd.concat([main_data_from_g, sub_data], 
                                                 axis=0, copy=False).reset_index(drop=True)
                    # 根据工作表名称决定数据写入的起始列
                    if sheet_name == "站外数据源":
                        start_col = "F"
                    elif sheet_name == "店铺成交数据源":
                        start_col = "H"
                    else:
                        start_col = "G"
                    new_sheet.range(f"{start_col}2").options(index=False, header=False).value = merged_data_from_g.values

                new_wb.save()
                new_wb.close()
                new_app.quit()

                total_time = time.time() - start_time
                self.update_status(f"由于原文件可能被锁定，已将结果保存到新文件：\n{save_path}\n处理耗时：{total_time:.2f}秒")
                messagebox.showinfo("成功", f"数据已成功合并并保存至新文件：\n{save_path}")
                self.progress_var.set(100)

        except Exception as e:
            logger.error(f"合并文件时出错: {str(e)}")
            try:
                wb.close()
                app.quit()
            except:
                pass
            messagebox.showerror("错误", f"合并文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 有足够的磁盘空间\n3. 有写入权限")

if __name__ == "__main__":
    ExcelMerger()