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
        self.debug_mode = None  # 先初始化为None
        
        # 工作表配置字典，存储工作表名称与起始列的映射关系
        self.sheet_config = {
            "全站营销": {"start_col": "G", "formula_end_col": "F", "date_col": None},
            "站内数据源": {"start_col": "G", "formula_end_col": "F", "date_col": None},
            "站外数据源": {"start_col": "F", "formula_end_col": "E", "date_col": None},
            "店铺成交数据源": {"start_col": "H", "formula_end_col": "F", "date_col": "G"}
        }
        
        # 文件名关键词映射，用于自动识别副表类型
        self.file_keywords = {
            "全站营销": "全站营销",
            "站内数据源": "日报数据",
            "站外数据源": "下单订单明细",
            "店铺成交数据源": "全部渠道"
        }
        
        self.setup_gui()
        # 注意：debug_mode已在setup_gui()中初始化，此处不需要再次初始化
        
    def adjust_formula_for_row(self, formula, template_row, target_row):
        """调整公式中的行引用，保持外部工作簿引用不变 - 优化版"""
        import re
        
        # 如果公式为空，直接返回
        if not formula:
            return formula
            
        # 行差值
        row_diff = target_row - template_row
        
        # 使用预编译的正则表达式提高性能
        # 匹配模式：[可能的外部引用]可能的工作表名!列字母行数字
        # 例如：[Book1.xlsx]Sheet1!A1 或 Sheet1!A1 或 A1
        if not hasattr(self, '_cell_ref_pattern'):
            self._cell_ref_pattern = re.compile(r'(\[.*?\])?([^\[\]!]+!)?([A-Za-z]+)([0-9]+)')
        
        def adjust_cell_ref(match):
            external_ref = match.group(1) or ''
            sheet_ref = match.group(2) or ''
            col_ref = match.group(3)
            row_ref = int(match.group(4))
            
            # 优化条件判断逻辑，减少分支
            if external_ref or sheet_ref:
                # 如果是对外部工作簿或其他工作表的引用，保持行号不变
                new_row_ref = row_ref
            elif row_ref == template_row:
                # 如果是对当前行的引用，则调整行号
                new_row_ref = target_row
            else:
                # 如果是对其他行的相对引用，则调整行号
                new_row_ref = row_ref + row_diff
            
            # 返回调整后的单元格引用
            return f"{external_ref}{sheet_ref}{col_ref}{new_row_ref}"
        
        # 使用预编译的正则表达式替换所有单元格引用
        adjusted_formula = self._cell_ref_pattern.sub(adjust_cell_ref, formula)
        
        return adjusted_formula

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
                # 清理内存
                process = psutil.Process(os.getpid())
                process.memory_info()

    def on_drop_main(self, event):
        """处理主表文件的拖放事件"""
        file_path = event.data
        try:
            if file_path.startswith('{'): # Windows 路径处理
                import json
                file_path = json.loads(file_path)['text']
            elif file_path.startswith('file://'): # macOS 路径处理
                file_path = file_path.replace('file://', '')
                # 处理URL编码的字符
                import urllib.parse
                file_path = urllib.parse.unquote(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"处理文件路径时出错：\n{str(e)}\n\n原始路径：{event.data}")
            return
            
        if os.path.isfile(file_path) and file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.et', '.ett')):
            self.load_main_file(file_path)  # 直接传递文件路径
        else:
            messagebox.showerror("错误", "请拖放有效的Excel文件！")

    def on_drop_sub(self, event, sheet_name):
        """处理副表文件的拖放事件"""
        file_path = event.data
        try:
            if file_path.startswith('{'): # Windows 路径处理
                import json
                file_path = json.loads(file_path)['text']
            elif file_path.startswith('file://'): # macOS 路径处理
                file_path = file_path.replace('file://', '')
                # 处理URL编码的字符
                import urllib.parse
                file_path = urllib.parse.unquote(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"处理文件路径时出错：\n{str(e)}\n\n原始路径：{event.data}")
            return
            
        if os.path.isfile(file_path):
            try:
                self.update_status(f"正在加载{sheet_name}的副表文件...")

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

                total_rows = len(self.sub_data[sheet_name])
                loaded_files = "\n".join([os.path.basename(f) for f in self.sub_files[sheet_name]])
                self.update_status(f"{sheet_name}副表文件加载成功\n已加载的文件：\n{loaded_files}\n总行数：{total_rows}")

            except Exception as e:
                messagebox.showerror("错误", f"加载文件 {os.path.basename(file_path)} 时出错：\n{str(e)}")
        else:
            messagebox.showerror("错误", "请拖放有效的Excel或CSV文件！")

    def update_status(self, message, level='info'):
        """更新状态信息
        
        参数:
            message: 状态信息内容
            level: 日志级别，可选值为 'debug', 'info', 'warning', 'error'
                  'debug': 仅在调试模式下显示
                  'info': 普通信息，默认级别
                  'warning': 警告信息，始终显示
                  'error': 错误信息，始终显示
        """
        # 根据日志级别和调试模式决定是否显示消息
        if level == 'debug' and not self.debug_mode.get():
            return  # 在非调试模式下不显示调试信息
        
        # 控制台输出（仅在调试模式下或非调试信息时）
        if self.debug_mode.get() or level != 'debug':
            print(f"[{level.upper()}] {message}")
        
        # 在GUI中显示消息
        # 启用文本框编辑
        self.status_text.configure(state='normal')
        
        # 根据级别添加前缀
        prefix = ""
        if level == 'warning':
            prefix = "警告: "
        elif level == 'error':
            prefix = "错误: "
        
        # 在文本框末尾添加新消息和两个换行（一个空行）
        self.status_text.insert(tk.END, prefix + message + "\n\n")
        # 滚动到最新位置
        self.status_text.see(tk.END)
        # 恢复只读模式
        self.status_text.configure(state='disabled')
        # 更新GUI
        self.root.update()

    def clear_all_files(self):
        """清理所有已加载的文件数据"""
        self.main_file = None
        self.sub_files = {}
        self.main_data = {}
        self.sub_data = {}
        print("已清理所有已加载的文件数据")
        self.update_status("已清理所有文件，请重新选择文件")
        
    def batch_load_sub_files(self):
        """批量加载副表文件，根据文件名自动识别类型"""
        # 添加警告过滤器
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        
        # 直接显示文件选择对话框，跳过提示信息
        file_paths = filedialog.askopenfilenames(filetypes=[
            ("All supported files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett *.csv"), 
            ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett"), 
            ("CSV files", "*.csv"), 
            ("All files", "*.*")])
        
        if not file_paths:  # 如果用户取消选择，直接返回
            return

        try:
            self.update_status("正在批量加载副表文件...")
            
            # 用于存储每个工作表的文件计数
            sheet_file_counts = {sheet: 0 for sheet in self.file_keywords.keys()}
            total_files = len(file_paths)
            recognized_files = 0
            unrecognized_files = []
            
            # 第一步：根据文件名分类文件
            categorized_files = {sheet: [] for sheet in self.file_keywords.keys()}
            
            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                recognized = False
                
                # 检查文件名中是否包含关键词
                for sheet_name, keyword in self.file_keywords.items():
                    if keyword in file_name:
                        categorized_files[sheet_name].append(file_path)
                        sheet_file_counts[sheet_name] += 1
                        recognized = True
                        recognized_files += 1
                        break
                
                if not recognized:
                    unrecognized_files.append(file_path)
            
            # 第二步：加载每个类别的文件
            for sheet_name, files in categorized_files.items():
                if not files:  # 如果该类别没有文件，跳过
                    continue
                    
                self.update_status(f"正在加载{sheet_name}的副表文件，共{len(files)}个...")
                
                # 初始化或重置该工作表的副表数据
                if sheet_name not in self.sub_data:
                    self.sub_data[sheet_name] = pd.DataFrame()
                    self.sub_files[sheet_name] = []
                
                # 批量处理文件，减少DataFrame合并次数
                batch_dfs = []
                error_count = 0
                
                for file_path in files:
                    try:
                        self.update_status(f"正在加载文件: {os.path.basename(file_path)}", level='debug')
                        
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

                        # 将DataFrame添加到批处理列表中
                        batch_dfs.append(df)
                        self.sub_files[sheet_name].append(file_path)
                        
                    except Exception as e:
                        error_count += 1
                        error_msg = f"加载文件 {os.path.basename(file_path)} 时出错：\n{str(e)}"
                        self.update_status(error_msg, level='error')
                        if error_count <= 3:  # 只显示前3个错误
                            messagebox.showerror("错误", error_msg)
                        continue
                
                # 一次性合并所有DataFrame
                if batch_dfs:
                    if sheet_name in self.sub_data and not self.sub_data[sheet_name].empty:
                        # 如果已有数据，添加到现有数据
                        all_dfs = [self.sub_data[sheet_name]] + batch_dfs
                        self.sub_data[sheet_name] = pd.concat(all_dfs, ignore_index=True)
                    else:
                        # 如果是首次加载，直接合并所有批次
                        self.sub_data[sheet_name] = pd.concat(batch_dfs, ignore_index=True)
                    
                    total_rows = len(self.sub_data[sheet_name])
                    self.update_status(f"{sheet_name}副表加载完成，共{len(files)-error_count}个文件，{total_rows}行数据")
            
            # 显示未识别的文件
            if unrecognized_files:
                unrecognized_count = len(unrecognized_files)
                if unrecognized_count <= 5:
                    unrecognized_names = "\n".join([os.path.basename(f) for f in unrecognized_files])
                    self.update_status(f"警告：以下{unrecognized_count}个文件未能识别类型：\n{unrecognized_names}", level='warning')
                else:
                    unrecognized_names = "\n".join([os.path.basename(f) for f in unrecognized_files[:5]]) + f"\n...等共{unrecognized_count}个文件"
                    self.update_status(f"警告：有{unrecognized_count}个文件未能识别类型，前5个为：\n{unrecognized_names}", level='warning')
            
            # 汇总加载结果
            summary = [f"{sheet}：{count}个文件" for sheet, count in sheet_file_counts.items() if count > 0]
            if summary:
                self.update_status(f"批量导入完成！成功识别并加载{recognized_files}个文件:\n" + "\n".join(summary))
            else:
                self.update_status("批量导入完成，但未能识别任何文件类型。请检查文件名是否包含正确的关键词。")
                
        except Exception as e:
            messagebox.showerror("错误", f"批量加载副表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏\n4. CSV文件编码格式正确")

    def setup_gui(self):
        """设置GUI界面"""
        self.root = TkinterDnD.Tk()
        self.root.title("Excel文件合并工具 v1.4")
        self.root.geometry("500x800")
        self.root.minsize(500, 800)
        
        # 初始化debug_mode变量
        if self.debug_mode is None:
            self.debug_mode = tk.BooleanVar(value=False)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 设置样式
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 10))
        style.configure("TLabel", font=("Arial", 10))
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"))
        style.configure("TCheckbutton", font=("Arial", 10))

        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        # 主表文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="主表文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=10)

        main_file_frame = ttk.Frame(file_frame)
        main_file_frame.pack(fill=tk.X, pady=5)
        
        # 左侧按钮区域
        button_area = ttk.Frame(main_file_frame)
        button_area.pack(side=tk.LEFT, padx=10)
        select_main_btn = ttk.Button(button_area, text="选择主表文件", width=15, 
                                   command=lambda: self.load_main_file(None))
        select_main_btn.pack()
        
        # 右侧拖拽区域
        main_drop_frame = ttk.LabelFrame(main_file_frame, text="拖拽区域")
        main_drop_frame.pack(side=tk.RIGHT, padx=10, fill=tk.X, expand=True)
        main_drop_frame.drop_target_register("DND_Files")
        main_drop_frame.dnd_bind('<<Drop>>', self.on_drop_main)
        drop_label = ttk.Label(main_drop_frame, text="拖拽文件到这里", anchor="center")
        drop_label.pack(pady=10, fill=tk.X)

        # 副表文件选择区域
        sub_files_frame = ttk.LabelFrame(main_frame, text="副表文件选择", padding=10)
        sub_files_frame.pack(fill=tk.X, pady=10)

        # 初始化复选框变量
        self.merge_marketing = tk.BooleanVar(value=True)
        self.merge_internal = tk.BooleanVar(value=True)
        self.merge_external = tk.BooleanVar(value=True)
        self.merge_shop = tk.BooleanVar(value=True)

        # 创建副表选择的函数
        def create_sub_file_frame(parent, text, var, sheet_name):
            frame = ttk.Frame(parent)
            frame.pack(fill=tk.X, pady=5)
            
            # 左侧控制区域
            control_area = ttk.Frame(frame)
            control_area.pack(side=tk.LEFT, padx=10, anchor="w")
            
            # 复选框和按钮水平排列
            ttk.Checkbutton(control_area, text=text, variable=var).pack(side=tk.LEFT, padx=(0, 10))
            ttk.Button(control_area, text="选择副表", width=12,
                      command=lambda: self.load_sub_file(sheet_name)).pack(side=tk.LEFT)
            
            # 右侧拖拽区域
            drop_frame = ttk.LabelFrame(frame, text="拖拽文件到这里")
            drop_frame.pack(side=tk.RIGHT, padx=10, fill=tk.X, expand=True)
            drop_frame.drop_target_register("DND_Files")
            drop_frame.dnd_bind('<<Drop>>', lambda e: self.on_drop_sub(e, sheet_name))
            ttk.Label(drop_frame, text=f"{sheet_name}副表", anchor="center").pack(pady=8, fill=tk.X)

        # 使用函数创建各个副表框架
        create_sub_file_frame(sub_files_frame, "全站营销", self.merge_marketing, "全站营销")
        create_sub_file_frame(sub_files_frame, "站内数据源", self.merge_internal, "站内数据源")
        create_sub_file_frame(sub_files_frame, "站外数据源", self.merge_external, "站外数据源")
        create_sub_file_frame(sub_files_frame, "店铺成交数据源", self.merge_shop, "店铺成交数据源")

        # 操作按钮区域
        button_frame = ttk.LabelFrame(main_frame, text="操作", padding=10)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 按钮居中显示
        buttons_container = ttk.Frame(button_frame)
        buttons_container.pack(pady=5, fill=tk.X)
        
        # 批量导入副表按钮
        batch_import_button = ttk.Button(buttons_container, text="批量导入副表", width=15, command=self.batch_load_sub_files)
        batch_import_button.pack(side=tk.LEFT, padx=10, expand=True)
        
        # 合并按钮
        merge_button = ttk.Button(buttons_container, text="合并文件", width=15, command=self.merge_files)
        merge_button.pack(side=tk.LEFT, padx=10, expand=True)

        # 清理按钮
        clear_button = ttk.Button(buttons_container, text="清理所有文件", width=15, command=self.clear_all_files)
        clear_button.pack(side=tk.RIGHT, padx=10, expand=True)

        # 状态信息区域
        status_frame = ttk.LabelFrame(main_frame, text="状态信息", padding=10)
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 添加调试模式复选框
        debug_frame = ttk.Frame(status_frame)
        debug_frame.pack(fill=tk.X, padx=10, pady=2, anchor=tk.W)
        debug_check = ttk.Checkbutton(debug_frame, text="调试模式（显示详细错误信息）", variable=self.debug_mode)
        debug_check.pack(side=tk.LEFT)

        # 状态标签
        self.status_label = ttk.Label(status_frame, text="请选择文件", wraplength=400, justify=tk.LEFT)
        self.status_label.pack(pady=5, fill=tk.X, padx=10)

        # 状态文本框，支持滚动显示
        text_frame = ttk.Frame(status_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 文本框
        self.status_text = tk.Text(text_frame, wrap=tk.WORD, height=10, width=80, 
                                  font=("Arial", 9), bd=1, relief=tk.SOLID)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
        
        # 设置为只读模式
        self.status_text.configure(state='disabled')

        self.root.mainloop()

    def load_main_file(self, file_path=None):
        """加载主表文件"""
        # 添加警告过滤器
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        
        if file_path is None:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.et *.ett")])
            if not file_path:  # 如果用户取消选择，直接返回
                return
        elif isinstance(file_path, tuple):
            file_path = file_path[0]
        
        try:
            self.update_status("正在加载主表文件...")
            
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
                for sheet_name in required_sheets:
                    self.main_data[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)

            except Exception as e:
                # 如果pandas读取失败，尝试使用xlwings读取
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
                    for sheet_name in required_sheets:
                        self.main_data[sheet_name] = wb.sheets[sheet_name].used_range.options(pd.DataFrame, index=False).value

                    wb.close()
                    app.quit()

                except Exception as e2:
                    messagebox.showerror("错误", f"无法读取主表文件，请确保文件格式正确。\n错误详情：\n{str(e)}\n{str(e2)}")
                    return

            self.main_file = file_path
            sheet_info = "\n".join([f"{sheet}：{len(data)}行" for sheet, data in self.main_data.items()])
            self.update_status(f"主表文件加载成功\n文件路径：{file_path}\n{sheet_info}")
        except Exception as e:
            messagebox.showerror("错误", f"加载主表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏")

    def load_sub_file(self, sheet_name):
        """加载指定工作表的副表文件"""
        # 添加警告过滤器
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        
        # 只有当通过按钮点击时才显示文件选择对话框
        file_paths = filedialog.askopenfilenames(filetypes=[
            ("All supported files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett *.csv"), 
            ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.et *.ett"), 
            ("CSV files", "*.csv"), 
            ("All files", "*.*")])
        
        if not file_paths:  # 如果用户取消选择，直接返回
            return

        try:
            self.update_status(f"正在加载{sheet_name}的副表文件...", level='info')
            
            # 初始化或重置该工作表的副表数据
            if sheet_name not in self.sub_data:
                self.sub_data[sheet_name] = pd.DataFrame()
                self.sub_files[sheet_name] = []

            # 批量处理文件，减少DataFrame合并次数
            batch_dfs = []
            file_count = len(file_paths)
            loaded_count = 0
            error_count = 0
            total_rows = 0
            
            # 循环处理每个选中的文件
            for i, file_path in enumerate(file_paths):
                try:
                    # 只在调试模式下显示每个文件的加载信息
                    self.update_status(f"正在加载文件({i+1}/{file_count}): {os.path.basename(file_path)}", level='debug')
                    
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

                    # 将DataFrame添加到批处理列表中，而不是每次都合并
                    batch_dfs.append(df)
                    self.sub_files[sheet_name].append(file_path)
                    loaded_count += 1
                    total_rows += len(df)

                except Exception as e:
                    error_count += 1
                    error_msg = f"加载文件 {os.path.basename(file_path)} 时出错：\n{str(e)}"
                    self.update_status(error_msg, level='error')
                    # 只在错误较少时显示错误对话框，避免大量文件时弹出过多对话框
                    if error_count <= 3:
                        messagebox.showerror("错误", error_msg)
                    continue

            # 一次性合并所有DataFrame，提高效率
            if batch_dfs:
                if sheet_name in self.sub_data and not self.sub_data[sheet_name].empty:
                    # 如果已有数据，添加到现有数据
                    all_dfs = [self.sub_data[sheet_name]] + batch_dfs
                    self.sub_data[sheet_name] = pd.concat(all_dfs, ignore_index=True)
                else:
                    # 如果是首次加载，直接合并所有批次
                    self.sub_data[sheet_name] = pd.concat(batch_dfs, ignore_index=True)
            
            # 汇总加载结果
            if loaded_count > 0:
                # 只显示前5个文件名，如果超过5个则显示省略号
                file_names = [os.path.basename(f) for f in self.sub_files[sheet_name][-loaded_count:]]
                if len(file_names) > 5:
                    displayed_files = "\n".join(file_names[:5]) + f"\n...等共{len(file_names)}个文件"
                else:
                    displayed_files = "\n".join(file_names)
                
                self.update_status(f"{sheet_name}副表文件加载成功: 共{loaded_count}个文件，{total_rows}行数据" + 
                                  (f"\n已加载的文件：\n{displayed_files}" if self.debug_mode.get() else ""), 
                                  level='info')
            
            # 如果有错误，显示汇总信息
            if error_count > 0:
                self.update_status(f"警告：{error_count}个文件加载失败，已跳过这些文件", level='warning')

        except Exception as e:
            messagebox.showerror("错误", f"加载副表文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 文件格式正确\n3. 文件未损坏\n4. CSV文件编码格式正确")

    def safe_apply_formula(self, sheet, range_str, formulas, retry_on_error=True, max_retries=3):
        """安全地应用公式，处理可能的外部引用错误"""
        retry_count = 0
        original_batch_size = len(formulas) if isinstance(formulas, list) else 1
        
        while retry_count < max_retries:
            try:
                # 尝试应用公式
                sheet.range(range_str).formula = formulas
                return True, "成功"
            except Exception as e:
                error_str = str(e).lower()
                retry_count += 1
                
                # 处理特定类型的错误
                if "apple event timed out" in error_str or "oserror: -1712" in error_str:
                    if not retry_on_error or retry_count >= max_retries:
                        return False, f"Apple event超时错误: {str(e)[:100]}"
                    
                    # 如果是批量操作，尝试减小批量大小
                    if isinstance(formulas, list) and len(formulas) > 1:
                        # 分割批次
                        mid = len(formulas) // 2
                        range_parts = range_str.split(":")
                        if len(range_parts) == 2:
                            start_cell, end_cell = range_parts
                            col = start_cell[0]
                            start_row = int(start_cell[1:])
                            end_row = int(end_cell[1:])
                            mid_row = start_row + mid - 1
                            
                            # 递归处理前半部分
                            range_str1 = f"{col}{start_row}:{col}{mid_row}"
                            success1, _ = self.safe_apply_formula(sheet, range_str1, formulas[:mid], True, max_retries-1)
                            
                            # 递归处理后半部分
                            range_str2 = f"{col}{mid_row+1}:{col}{end_row}"
                            success2, _ = self.safe_apply_formula(sheet, range_str2, formulas[mid:], True, max_retries-1)
                            
                            return success1 or success2, "部分成功"
                elif "找不到" in error_str or "not found" in error_str or "cannot find" in error_str:
                    return False, f"找不到引用的外部工作簿: {str(e)[:100]}"
                else:
                    if not retry_on_error or retry_count >= max_retries:
                        return False, f"未知错误: {str(e)[:100]}"
        
        return False, "达到最大重试次数"
    
    def fill_sheet_formula(self, wb, target_sheet_name, end_column='F'):
        """在指定工作表的A到指定结束列填充公式 - 通用版
        
        此函数实现了Excel工作表的"前置填充"功能，可以自动将模板行(通常是第2行)的公式复制并应用到后续所有数据行。
        函数会智能处理外部工作簿引用的公式，并针对不同类型的公式采用不同的填充策略以提高性能和成功率。
        
        主要特点：
        1. 支持模糊匹配工作表名称，提高用户体验
        2. 智能识别外部工作簿引用，并采用批处理方式提高性能
        3. 针对外部引用公式失败时采用单行回退策略，确保最大程度完成任务
        4. 多线程并行处理多列，显著提高处理速度
        5. 内存优化，避免处理大量数据时内存溢出
        6. 详细的错误处理和用户反馈
        
        参数:
            wb: Excel工作簿对象，通过xlwings打开的工作簿
            target_sheet_name: 目标工作表名称，支持模糊匹配
            end_column: 结束列，默认为'F'，可以根据工作表类型设置为'E'或其他列
                       例如：'E'表示处理A到E列的公式
        
        返回:
            无直接返回值，处理结果通过update_status方法反馈给用户界面
        
        异常:
            处理各种可能的异常，包括工作表不存在、公式错误、外部引用不可访问等
            对常见错误提供友好的错误信息和建议
        """
        try:
            import re
            import threading
            import concurrent.futures
            from functools import partial
            import gc
            
            # 减少日志输出，只保留关键信息
            self.update_status(f"开始处理{target_sheet_name}工作表A到{end_column}列公式填充...")
            
            # 获取所有工作表名称
            sheet_names = [sheet.name for sheet in wb.sheets]
            
            # 优化工作表查找逻辑，使用更高效的方法
            sheet_found = False
            actual_sheet_name = ""
            
            # 使用字典查找提高效率
            sheet_dict = {name.lower().strip(): name for name in sheet_names}
            
            # 首先尝试精确匹配
            if target_sheet_name in sheet_names:
                sheet_found = True
                actual_sheet_name = target_sheet_name
            # 然后尝试不区分大小写的匹配
            elif target_sheet_name.lower().strip() in sheet_dict:
                sheet_found = True
                actual_sheet_name = sheet_dict[target_sheet_name.lower().strip()]
                self.update_status(f"找到近似匹配的工作表: '{actual_sheet_name}'")
            # 最后尝试部分匹配
            else:
                for name in sheet_names:
                    if target_sheet_name in name or name in target_sheet_name:
                        sheet_found = True
                        actual_sheet_name = name
                        self.update_status(f"找到部分匹配的工作表: '{actual_sheet_name}'")
                        break
            
            if not sheet_found:
                self.update_status(f"未找到{target_sheet_name}工作表，跳过公式填充")
                return

            # 使用找到的实际工作表名称
            sheet = wb.sheets[actual_sheet_name]
            
            # 使用used_range来获取整个工作表的使用范围，确保包含新添加的数据
            used_range = sheet.used_range
            last_row = used_range.last_cell.row
            self.update_status(f"工作表最后一行为第{last_row}行")
            
            # 定义需要处理的列，从A列到指定的结束列
            columns = [chr(ord('A') + i) for i in range(ord(end_column) - ord('A') + 1)]
            
            # 从第3行开始填充到最后一行
            if last_row > 2:  # 确保有数据行需要填充
                total_start_time = time.time()
                
                # 预编译正则表达式以提高性能
                external_ref_pattern = re.compile(r'\[.*?\].*?!')
                
                # 创建一个缓存字典，用于存储已经计算过的公式调整结果
                formula_cache = {}
                
                # 定义一个函数来处理单列的公式填充
                def process_column(col):
                    col_start_time = time.time()
                    
                    # 获取第2行的公式（作为模板）
                    template_cell = sheet.range(f'{col}2')
                    template_formula = template_cell.formula
                    
                    # 如果当前列没有公式，尝试检查第1行
                    if not template_formula:
                        template_formula = sheet.range(f'{col}1').formula
                        if not template_formula:
                            return f"{col}列没有可用的公式模板，已跳过"
                    
                    # 检查是否包含外部工作簿引用
                    has_external_ref = external_ref_pattern.search(template_formula) is not None
                    
                    # 根据是否包含外部引用调整批处理大小
                    # 增加批处理大小，但对外部引用保持较小的批量以确保稳定性
                    batch_size = 400 if has_external_ref else 3000  # 外部引用使用较小的批处理大小
                    total_rows = last_row - 2
                    
                    # 添加重试次数限制，避免无限循环
                    max_retries = 3
                    
                    if has_external_ref:
                        # 对于包含外部引用的公式，需要为每行生成调整后的公式
                        batches = (total_rows + batch_size - 1) // batch_size  # 向上取整
                        
                        # 提取外部工作簿路径，用于错误处理
                        external_workbook_match = re.search(r'\[([^\]]+)\]', template_formula)
                        external_workbook_name = external_workbook_match.group(1) if external_workbook_match else "未知外部工作簿"
                        
                        # 记录成功和失败的行数
                        success_count = 0
                        fail_count = 0
                        
                        try:
                            for batch in range(batches):
                                start_idx = 3 + batch * batch_size
                                end_idx = min(start_idx + batch_size - 1, last_row)
                                
                                # 使用列表推导式优化公式生成
                                formulas = []
                                for row in range(start_idx, end_idx + 1):
                                    # 使用缓存减少重复计算
                                    cache_key = (template_formula, 2, row)
                                    if cache_key not in formula_cache:
                                        formula_cache[cache_key] = self.adjust_formula_for_row(template_formula, 2, row)
                                    formulas.append([formula_cache[cache_key]])
                                
                                # 使用安全应用公式函数
                                range_str = f"{col}{start_idx}:{col}{end_idx}"
                                success, error_msg = self.safe_apply_formula(sheet, range_str, formulas, True, max_retries)
                                
                                if success:
                                    success_count += (end_idx - start_idx + 1)
                                else:
                                    # 如果批量应用失败，尝试使用较小的批次或单行应用
                                    self.update_status(f"警告：{col}列批量设置公式时出错({error_msg})，尝试单行处理...", level='warning')
                                    
                                    # 增加更新频率，大幅减少UI更新次数
                                    update_frequency = 100
                                    
                                    # 逐行设置公式，即使部分失败也继续处理
                                    for i, row in enumerate(range(start_idx, end_idx + 1)):
                                        try:
                                            # 单行设置公式
                                            formula = formulas[i][0]
                                            single_range = f"{col}{row}"
                                            single_success, _ = self.safe_apply_formula(sheet, single_range, [formula], False, 1)
                                            
                                            if single_success:
                                                success_count += 1
                                            else:
                                                fail_count += 1
                                                
                                            # 大幅减少状态更新频率，只在调试模式下更新
                                            if i % update_frequency == 0 and i > 0:
                                                # 每处理一定数量的行就更新一下状态
                                                self.update_status(f"正在处理{col}列，已成功{success_count}行，失败{fail_count}行...", level='debug')
                                        except Exception:
                                            # 如果单行设置也失败，记录错误但继续处理
                                            fail_count += 1
                                
                                # 主动清理内存
                                if len(formula_cache) > 5000:  # 降低缓存阈值，更频繁地清理内存
                                    formula_cache.clear()
                                    gc.collect()
                                    
                            # 批次处理完成后，报告结果
                            if fail_count > 0:
                                self.update_status(f"{col}列公式填充完成，成功{success_count}行，失败{fail_count}行")
                                if external_workbook_match:
                                    self.update_status(f"提示：失败可能是因为找不到外部工作簿 '{external_workbook_name}'，请确保该文件存在且可访问")
                            
                        except Exception as outer_exc:
                            # 处理整个批处理过程中的异常
                            return f"{col}列公式填充过程中出错: {str(outer_exc)[:150]}，已尽可能填充部分数据(成功{success_count}行，失败{fail_count}行)"
                    else:
                        # 对于不包含外部引用的普通公式，也需要为每行生成调整后的公式
                        # 使用批处理方式提高效率
                        batches = (total_rows + batch_size - 1) // batch_size  # 向上取整
                        
                        for batch in range(batches):
                            start_idx = 3 + batch * batch_size
                            end_idx = min(start_idx + batch_size - 1, last_row)
                            
                            # 使用列表推导式优化公式生成
                            formulas = []
                            for row in range(start_idx, end_idx + 1):
                                # 使用缓存减少重复计算
                                cache_key = (template_formula, 2, row)
                                if cache_key not in formula_cache:
                                    formula_cache[cache_key] = self.adjust_formula_for_row(template_formula, 2, row)
                                formulas.append([formula_cache[cache_key]])
                            
                            # 应用公式
                            range_str = f"{col}{start_idx}:{col}{end_idx}"
                            sheet.range(range_str).formula = formulas
                    
                    col_end_time = time.time()
                    return f"已成功在{target_sheet_name}工作表{col}列填充公式，共处理{last_row-2}行，耗时{col_end_time-col_start_time:.2f}秒"
                
                # 使用线程池并行处理多列
                results = []
                self.update_status(f"开始并行处理{len(columns)}列公式填充...", level='info')
                # 导入multiprocessing模块以获取CPU核心数
                import multiprocessing
                # 增加线程数量，使用CPU核心数或8个线程（取较小值）
                cpu_count = multiprocessing.cpu_count()
                thread_count = min(8, cpu_count, len(columns))
                self.update_status(f"使用{thread_count}个线程并行处理{len(columns)}列公式填充...", level='info')
                with concurrent.futures.ThreadPoolExecutor(max_workers=thread_count) as executor:
                    # 提交所有列的处理任务
                    future_to_col = {executor.submit(process_column, col): col for col in columns}
                    
                    # 收集结果
                    for future in concurrent.futures.as_completed(future_to_col):
                        col = future_to_col[future]
                        try:
                            # 增加超时处理，防止某一列长时间阻塞
                            result = future.result(timeout=300)  # 设置5分钟超时
                            results.append(result)
                            # 只在调试模式下输出每列完成的详细信息
                            self.update_status(f"已完成{col}列公式填充", level='debug')
                        except concurrent.futures.TimeoutError:
                            # 处理超时情况
                            results.append(f"{col}列处理超时，可能是外部引用导致，已跳过该列")
                            self.update_status(f"警告：{col}列处理时间过长，已跳过。请检查外部引用是否可访问。", level='warning')
                        except Exception as exc:
                            # 处理其他异常
                            error_str = str(exc).lower()
                            if "apple event timed out" in error_str or "oserror: -1712" in error_str:
                                results.append(f"{col}列处理时出错: Apple event超时，可能是外部工作簿引用问题")
                                self.update_status(f"警告：{col}列处理时出现Apple event超时，已跳过该列。请确保外部引用的工作簿可访问。", level='warning')
                            elif "找不到" in error_str or "not found" in error_str or "cannot find" in error_str:
                                results.append(f"{col}列处理时出错: 找不到引用的外部工作簿")
                                self.update_status(f"警告：{col}列处理时找不到引用的外部工作簿，已跳过该列。", level='warning')
                            else:
                                results.append(f"{col}列处理时出错: {exc}")
                                self.update_status(f"处理{col}列时出错: {exc}", level='error')
                
                # 清理缓存和内存
                formula_cache.clear()
                gc.collect()
                
                total_end_time = time.time()
                
                # 只输出关键结果信息
                success_count = 0
                error_count = 0
                for result in results:
                    if "出错" in result or "错误" in result:
                        self.update_status(result, level='warning')
                        error_count += 1
                    elif "已跳过" in result:
                        self.update_status(result, level='debug')
                    else:
                        success_count += 1
                        # 只在调试模式下输出详细的成功信息
                        self.update_status(result, level='debug')
                
                # 输出汇总信息
                self.update_status(f"公式填充完成: 成功{success_count}列，失败{error_count}列", level='info')
                
                self.update_status(f"已完成{target_sheet_name}工作表A到F列的公式填充，总耗时{total_end_time-total_start_time:.2f}秒")
            else:
                self.update_status(f"{target_sheet_name}工作表没有足够的数据行需要填充公式")

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            error_str = str(e).lower()
            
            # 针对特定错误类型提供更友好的错误信息
            if "apple event timed out" in error_str or "oserror: -1712" in error_str:
                self.update_status(f"填充{target_sheet_name}工作表公式时出现超时错误，这可能是由于外部工作簿引用导致的。", level='error')
                self.update_status(f"建议：请确保所有引用的外部工作簿都可以访问，或考虑减少批处理大小。", level='info')
            elif "找不到" in error_str or "not found" in error_str or "cannot find" in error_str:
                self.update_status(f"填充{target_sheet_name}工作表公式时出错：找不到引用的外部工作簿。", level='error')
                self.update_status(f"建议：请确保所有引用的外部工作簿都在正确的位置，或更新公式中的引用路径。", level='info')
            else:
                self.update_status(f"填充{target_sheet_name}工作表公式时出错：{str(e)}", level='error')
                
            # 仅在调试模式下显示详细错误信息
            self.update_status(f"详细错误信息：\n{error_details}", level='debug')
            if not self.debug_mode.get():
                self.update_status(f"如需查看详细错误信息，请启用调试模式。", level='info')

    def merge_files(self):
        """工作表更新功能 - 将副表数据合并填充到主表
        
        此功能实现了将副表数据智能合并到主表的核心处理逻辑，主要包括以下步骤：
        1. 数据准备：打开主表文件，保留公式和格式
        2. 数据合并：根据工作表类型确定起始列，将副表数据追加到主表
        3. 特殊处理：针对店铺成交数据源工作表，从文件名提取日期信息
        4. 数据清理：处理空行、空列，确保数据完整性
        5. 公式填充：合并后自动填充公式，保证数据计算的连续性
        6. 结果保存：优先保存到原文件，失败时创建新文件
        
        支持的工作表类型：
        - 全站营销：数据从G列开始，公式填充A-F列
        - 站内数据源：数据从G列开始，公式填充A-F列
        - 站外数据源：数据从F列开始，公式填充A-E列
        - 店铺成交数据源：数据从H列开始，G列填充日期，公式填充A-F列
        """
        if not self.main_data or not self.sub_data:
            messagebox.showerror("错误", "请先加载主表和副表文件！")
            return

        if not any([self.merge_marketing.get(), self.merge_internal.get(), self.merge_external.get(), self.merge_shop.get()]):
            messagebox.showerror("错误", "请至少选择一个要合并的工作表！")
            return

        try:
            self.update_status("正在处理数据...")
            start_time = time.time()

            # 使用xlwings打开主表文件以保持公式和格式
            # 注意：这是工作表更新功能的关键步骤，使用xlwings而非pandas是为了保留Excel公式
            app = xw.App(visible=False)
            wb = app.books.open(self.main_file)

            # 工作表更新功能的核心循环：遍历所有副表数据并合并到对应的主表工作表
            selected_sheets = []
            for sheet_name, sub_data in self.sub_data.items():
                # 检查工作表是否被用户选中进行合并
                # 工作表更新功能支持选择性合并，用户可以决定哪些工作表需要更新
                if sheet_name == "全站营销" and not self.merge_marketing.get():
                    continue
                if sheet_name == "站内数据源" and not self.merge_internal.get():
                    continue
                if sheet_name == "站外数据源" and not self.merge_external.get():
                    continue
                if sheet_name == "店铺成交数据源" and not self.merge_shop.get():
                    continue
                
                # 确保主表中存在对应的工作表
                if sheet_name not in self.main_data:
                    continue
                
                selected_sheets.append(sheet_name)

                # 获取当前工作表对象和使用范围
                self.update_status(f"开始处理 {sheet_name} 工作表...", level='info')
                main_sheet = wb.sheets[sheet_name]
                main_range = main_sheet.used_range
                
                # 【工作表更新功能】根据配置字典确定数据起始列偏移
                # 不同类型的工作表有不同的数据结构，前面的列通常包含公式或标识信息
                start_col = self.sheet_config[sheet_name]["start_col"]
                # 将列字母转换为偏移量（A=0, B=1, ...）
                start_col_offset = ord(start_col) - ord('A')
                    
                # 从指定列偏移开始提取主表现有数据
                raw_data = main_range.offset(0, start_col_offset).options(pd.DataFrame, index=False, expand='table').value
                main_data_from_g = pd.DataFrame(raw_data)
                original_columns_count = len(main_data_from_g.columns)  # 记录原始列数，用于后续列数匹配检查

                # 【工作表更新功能-特殊处理】检查是否需要从文件名中提取日期并填充到指定列
                # 这是工作表更新功能的一个特殊处理逻辑，针对需要日期信息提取的工作表
                date_col = self.sheet_config[sheet_name]["date_col"]
                if date_col is not None:
                    # 获取主表当前数据的行数和副表的行数，用于确定填充位置
                    main_data_rows = len(main_data_from_g)
                    sub_data_rows = len(sub_data)
                    # 计算G列需要填充的起始行和结束行
                    # 从主表数据后的下一行开始，保留一行空行作为分隔
                    start_row = main_data_rows + 2
                    
                    # 从每个文件名中提取日期信息并填充对应的行
                    # 日期格式预期为文件名中的8位数字，如：_20230101_
                    import re
                    g_col_data = []
                    current_row = 0
                    
                    # 【工作表更新功能-日期提取】遍历每个副表文件，提取日期信息
                    for file_path in self.sub_files[sheet_name]:
                        # 读取当前文件的数据，支持CSV和Excel格式
                        if file_path.lower().endswith('.csv'):
                            try:
                                # 首先尝试使用默认编码读取CSV
                                df = pd.read_csv(file_path)
                            except UnicodeDecodeError:
                                # 如果出现编码错误，自动检测编码并重试
                                encoding = self.detect_encoding(file_path)
                                df = pd.read_csv(file_path, encoding=encoding)
                        else:
                            # 使用通用Excel加载函数处理各种Excel格式
                            df = self.load_excel_file(file_path)
                        
                        # 【工作表更新功能-智能日期提取】从文件名中提取日期
                        # 预期格式：文件名中包含_YYYYMMDD_格式的日期
                        file_name = os.path.basename(file_path)
                        date_match = re.search(r'_([0-9]{8})_', file_name)
                        if date_match:
                            # 成功提取到日期，转换为整数格式
                            date_value = int(date_match.group(1))
                        else:
                            print(f"警告：未能从文件名中提取到日期信息: {file_name}，使用固定值'error'")
                            self.update_status(f"警告：未能从文件名中提取到日期信息: {file_name}，使用固定值'error'")
                            date_value = "error"
                        
                        # 为当前文件的每一行数据添加对应的日期
                        file_rows = len(df)
                        g_col_data.extend([[date_value] for _ in range(file_rows)])
                        current_row += file_rows
                    
                    # 填充日期列数据
                    end_row = start_row + len(g_col_data) - 1
                    main_sheet.range(f"{date_col}{start_row}:{date_col}{end_row}").value = g_col_data

                # 【工作表更新功能-数据清理】优化列名处理
                # 处理可能存在的空列名，确保所有列都有有效名称
                if main_data_from_g.columns.isnull().any():
                    main_data_from_g.columns = [f"Column_{i+1}" if pd.isnull(col) else col 
                                               for i, col in enumerate(main_data_from_g.columns)]

                # 【工作表更新功能-数据清理】优化空列处理
                # 确保只有当整列（包括表头）都为空时才删除，避免误删有用数据
                empty_columns = main_data_from_g.isna().all() & main_data_from_g.columns.isna()
                if empty_columns.any():
                    # 记录非空列的索引，用于筛选保留的列
                    valid_columns = ~empty_columns
                    # 确保至少保留一列数据，防止数据结构丢失
                    if valid_columns.any():
                        main_data_from_g = main_data_from_g.loc[:, valid_columns]
                    else:
                        # 如果所有列都是空的，保持原样并发出警告
                        print(f"警告：{sheet_name}工作表所有列都是空的，保持原始数据结构")
                        self.update_status(f"警告：{sheet_name}工作表所有列都是空的，保持原始数据结构")

                # 优化空行处理，只处理连续的空行
                empty_rows = main_data_from_g.isna().all(axis=1)
                if empty_rows.any():
                    # 找到第一个完全空行的索引
                    first_empty_row_idx = empty_rows.idxmax()
                    # 检查是否所有后续行都是空行
                    if empty_rows[first_empty_row_idx:].all():
                        main_data_from_g = main_data_from_g.iloc[:first_empty_row_idx]
                        print(f"信息：从第{first_empty_row_idx}行开始检测到连续空行，已截断数据")
                        self.update_status(f"信息：从第{first_empty_row_idx}行开始检测到连续空行，已截断数据")

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
                # 使用配置字典获取起始列
                start_col = self.sheet_config[sheet_name]["start_col"]
                
                # 获取指定列的所有值
                col_values = main_sheet.range(f"{start_col}1").expand('down').value
                
                # 找到最后一个非空单元格的位置
                last_non_empty_row = 1  # 默认从第一行开始
                for i, value in enumerate(col_values, 1):
                    if value is not None and str(value).strip() != "":
                        last_non_empty_row = i
                
                # 在最后一个非空单元格后追加新数据
                append_start_row = last_non_empty_row + 1

                # 根据配置字典决定数据写入的起始列
                start_col = self.sheet_config[sheet_name]["start_col"]
                main_sheet.range(f"{start_col}{append_start_row}").options(index=False, header=False).value = sub_data.values

                # 更新进度条
                self.update_status(f"已完成{sheet_name}工作表的数据合并")

            try:
                # 保存文件前填充工作表的公式，使用配置字典中的formula_end_col值
                # 全站营销工作表
                if self.merge_marketing.get() and "全站营销" in self.sub_files and self.sub_files["全站营销"]:
                    formula_end_col = self.sheet_config["全站营销"]["formula_end_col"]
                    self.update_status(f"准备填充全站营销工作表的A到{formula_end_col}列公式...")
                    self.fill_sheet_formula(wb, "全站营销", formula_end_col)
                    self.update_status(f"全站营销工作表A到{formula_end_col}列公式填充处理完成")
                else:
                    if not self.merge_marketing.get():
                        self.update_status("全站营销工作表未被选中，跳过公式填充")
                    else:
                        self.update_status("全站营销工作表未导入副表，跳过公式填充")
                
                # 站内数据源工作表
                if self.merge_internal.get() and "站内数据源" in self.sub_files and self.sub_files["站内数据源"]:
                    formula_end_col = self.sheet_config["站内数据源"]["formula_end_col"]
                    self.update_status(f"准备填充站内数据源工作表的A到{formula_end_col}列公式...")
                    self.fill_sheet_formula(wb, "站内数据源", formula_end_col)
                    self.update_status(f"站内数据源工作表A到{formula_end_col}列公式填充处理完成")
                else:
                    if not self.merge_internal.get():
                        self.update_status("站内数据源工作表未被选中，跳过公式填充")
                    else:
                        self.update_status("站内数据源工作表未导入副表，跳过公式填充")
                
                # 站外数据源工作表
                if self.merge_external.get() and "站外数据源" in self.sub_files and self.sub_files["站外数据源"]:
                    formula_end_col = self.sheet_config["站外数据源"]["formula_end_col"]
                    self.update_status(f"准备填充站外数据源工作表的A到{formula_end_col}列公式...")
                    self.fill_sheet_formula(wb, "站外数据源", formula_end_col)
                    self.update_status(f"站外数据源工作表A到{formula_end_col}列公式填充处理完成")
                else:
                    if not self.merge_external.get():
                        self.update_status("站外数据源工作表未被选中，跳过公式填充")
                    else:
                        self.update_status("站外数据源工作表未导入副表，跳过公式填充")
                
                # 店铺成交数据源工作表
                if self.merge_shop.get() and "店铺成交数据源" in self.sub_files and self.sub_files["店铺成交数据源"]:
                    formula_end_col = self.sheet_config["店铺成交数据源"]["formula_end_col"]
                    self.update_status(f"准备填充店铺成交数据源工作表的A到{formula_end_col}列公式...")
                    self.fill_sheet_formula(wb, "店铺成交数据源", formula_end_col)
                    self.update_status(f"店铺成交数据源工作表A到{formula_end_col}列公式填充处理完成")
                else:
                    if not self.merge_shop.get():
                        self.update_status("店铺成交数据源工作表未被选中，跳过公式填充")
                    else:
                        self.update_status("店铺成交数据源工作表未导入副表，跳过公式填充")

                # 保存文件
                self.update_status("正在保存文件...")
                wb.save()
                wb.close()
                app.quit()

                total_time = time.time() - start_time
                self.update_status(f"合并完成！\n数据已保存至原始文件：{self.main_file}\n处理耗时：{total_time:.2f}秒")
                messagebox.showinfo("成功", "数据已成功合并并保存至原始文件")

            except Exception as save_error:
                print(f"错误：保存原文件失败: {str(save_error)}")
                self.update_status(f"错误：保存原文件失败: {str(save_error)}")
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
                        # 根据配置字典决定起始列
                        start_col = self.sheet_config[sheet_name]["start_col"]
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

        except Exception as e:
            print(f"错误：合并文件时出错: {str(e)}")
            self.update_status(f"错误：合并文件时出错: {str(e)}")
            try:
                wb.close()
                app.quit()
            except:
                pass
            messagebox.showerror("错误", f"合并文件时出错：\n{str(e)}\n\n请确保：\n1. 文件未被其他程序占用\n2. 有足够的磁盘空间\n3. 有写入权限")

if __name__ == "__main__":
    ExcelMerger()