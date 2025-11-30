import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ExcelManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel文件管理器")
        self.root.geometry("1000x600")
        
        # 初始化变量
        self.file_path = ""
        self.df = None
        self.current_sheet = ""
        self.sheets = []
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建菜单栏
        self.create_menu()
        
        # 创建工具栏
        self.create_toolbar()
        
        # 创建工作表选择区域
        self.create_sheet_selection()
        
        # 创建数据编辑区域
        self.create_data_area()
        
        # 创建状态栏
        self.create_status_bar()
        
    def create_menu(self):
        """创建菜单栏"""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # 文件菜单
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="打开", command=self.open_file)
        file_menu.add_command(label="保存", command=self.save_file)
        file_menu.add_command(label="导入", command=self.import_file)
        file_menu.add_command(label="导出", command=self.export_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        # 编辑菜单
        edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="编辑", menu=edit_menu)
        edit_menu.add_command(label="新增记录", command=self.add_record)
        edit_menu.add_command(label="删除记录", command=self.delete_record)
        edit_menu.add_command(label="修改记录", command=self.modify_record)
        
        # 帮助菜单
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)
        
    def create_toolbar(self):
        """创建工具栏"""
        self.toolbar = ttk.Frame(self.main_frame)
        self.toolbar.pack(fill=tk.X, pady=5)
        
        # 创建工具栏按钮
        self.btn_open = ttk.Button(self.toolbar, text="打开", command=self.open_file)
        self.btn_open.pack(side=tk.LEFT, padx=2)
        
        self.btn_save = ttk.Button(self.toolbar, text="保存", command=self.save_file)
        self.btn_save.pack(side=tk.LEFT, padx=2)
        
        self.btn_add = ttk.Button(self.toolbar, text="新增", command=self.add_record)
        self.btn_add.pack(side=tk.LEFT, padx=2)
        
        self.btn_delete = ttk.Button(self.toolbar, text="删除", command=self.delete_record)
        self.btn_delete.pack(side=tk.LEFT, padx=2)
        
        self.btn_modify = ttk.Button(self.toolbar, text="修改", command=self.modify_record)
        self.btn_modify.pack(side=tk.LEFT, padx=2)
        
    def create_sheet_selection(self):
        """创建工作表选择区域"""
        self.sheet_frame = ttk.Frame(self.main_frame)
        self.sheet_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.sheet_frame, text="工作表：").pack(side=tk.LEFT)
        
        self.sheet_var = tk.StringVar()
        self.sheet_combobox = ttk.Combobox(self.sheet_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.on_sheet_change)
        
    def create_data_area(self):
        """创建数据编辑区域"""
        self.data_frame = ttk.Frame(self.main_frame)
        self.data_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建滚动条
        self.vsb = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL)
        self.hsb = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL)
        
        # 创建Treeview用于显示数据
        self.tree = ttk.Treeview(self.data_frame, columns=[], show="headings", 
                                yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)
        
        # 绑定双击事件，用于编辑单元格
        self.tree.bind("<Double-1>", self.on_cell_double_click)
        
        # 布局Treeview和滚动条
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # 创建编辑框，初始隐藏
        self.edit_entry = ttk.Entry(self.data_frame)
        self.edit_entry.bind("<FocusOut>", self.on_edit_focus_out)
        self.edit_entry.bind("<Return>", self.on_edit_return)
        self.edit_entry.bind("<Escape>", self.on_edit_escape)
        self.edit_entry.pack_forget()
        
        # 编辑状态变量
        self.editing_cell = None  # (item_id, column)
        
    def create_status_bar(self):
        """创建状态栏"""
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def open_file(self):
        """打开Excel文件"""
        filetypes = [("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        filepath = filedialog.askopenfilename(title="打开Excel文件", filetypes=filetypes)
        
        if filepath:
            try:
                self.status_var.set(f"正在打开文件: {filepath}")
                self.root.update_idletasks()
                
                # 读取Excel文件，获取所有工作表
                self.file_path = filepath
                xl = pd.ExcelFile(filepath)
                self.sheets = xl.sheet_names
                
                # 更新工作表选择下拉框
                self.sheet_combobox['values'] = self.sheets
                if self.sheets:
                    # 选择第一个工作表
                    self.sheet_var.set(self.sheets[0])
                    self.current_sheet = self.sheets[0]
                    self.load_sheet_data(self.current_sheet)
                
                self.status_var.set(f"已打开文件: {os.path.basename(filepath)}，共{len(self.sheets)}个工作表")
            except Exception as e:
                messagebox.showerror("错误", f"打开文件失败: {str(e)}")
                self.status_var.set("打开文件失败")
        
    def save_file(self):
        """保存Excel文件"""
        if self.df is not None and self.file_path:
            try:
                self.status_var.set("正在保存文件...")
                self.root.update_idletasks()
                
                # 创建ExcelWriter对象，保存所有工作表
                with pd.ExcelWriter(self.file_path, engine='openpyxl' if self.file_path.endswith('.xlsx') else 'xlwt') as writer:
                    # 先保存当前修改的工作表
                    self.df.to_excel(writer, sheet_name=self.current_sheet, index=False)
                    
                    # 保存其他未修改的工作表
                    xl = pd.ExcelFile(self.file_path)
                    for sheet in self.sheets:
                        if sheet != self.current_sheet:
                            df_other = pd.read_excel(self.file_path, sheet_name=sheet)
                            df_other.to_excel(writer, sheet_name=sheet, index=False)
                
                self.status_var.set(f"文件已保存: {os.path.basename(self.file_path)}")
                messagebox.showinfo("成功", "文件保存成功")
            except Exception as e:
                messagebox.showerror("错误", f"保存文件失败: {str(e)}")
                self.status_var.set("保存文件失败")
        else:
            messagebox.showwarning("警告", "没有可保存的数据")
        
    def import_file(self):
        """导入Excel文件"""
        # 打开文件选择对话框
        filetypes = [("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        filepath = filedialog.askopenfilename(title="导入Excel文件", filetypes=filetypes)
        
        if filepath:
            try:
                self.status_var.set(f"正在导入文件: {filepath}")
                self.root.update_idletasks()
                
                # 读取要导入的Excel文件
                xl = pd.ExcelFile(filepath)
                import_sheets = xl.sheet_names
                
                # 让用户选择要导入的工作表
                if import_sheets:
                    # 创建导入选项对话框
                    self.import_window = tk.Toplevel(self.root)
                    self.import_window.title("导入选项")
                    self.import_window.geometry("300x200")
                    self.import_window.resizable(False, False)
                    self.import_window.transient(self.root)
                    self.import_window.grab_set()
                    
                    # 工作表选择
                    ttk.Label(self.import_window, text="选择要导入的工作表:").pack(pady=10)
                    self.import_sheet_var = tk.StringVar()
                    self.import_sheet_var.set(import_sheets[0])
                    sheet_combo = ttk.Combobox(self.import_window, textvariable=self.import_sheet_var, values=import_sheets, state="readonly")
                    sheet_combo.pack(pady=5, padx=10, fill=tk.X)
                    
                    # 导入方式选择
                    ttk.Label(self.import_window, text="导入方式:").pack(pady=10)
                    self.import_mode_var = tk.StringVar()
                    self.import_mode_var.set("merge")
                    
                    mode_frame = ttk.Frame(self.import_window)
                    mode_frame.pack(pady=5)
                    
                    ttk.Radiobutton(mode_frame, text="合并到当前工作表", variable=self.import_mode_var, value="merge").pack(anchor=tk.W)
                    ttk.Radiobutton(mode_frame, text="创建新工作表", variable=self.import_mode_var, value="new").pack(anchor=tk.W)
                    
                    # 按钮框架
                    button_frame = ttk.Frame(self.import_window)
                    button_frame.pack(fill=tk.X, padx=10, pady=10)
                    
                    ttk.Button(button_frame, text="确定", command=lambda: self.perform_import(filepath)).pack(side=tk.RIGHT, padx=5)
                    ttk.Button(button_frame, text="取消", command=self.import_window.destroy).pack(side=tk.RIGHT, padx=5)
                
            except Exception as e:
                messagebox.showerror("错误", f"导入文件失败: {str(e)}")
                self.status_var.set("导入文件失败")
    
    def perform_import(self, filepath):
        """执行导入操作"""
        try:
            # 获取选择的工作表和导入方式
            sheet_name = self.import_sheet_var.get()
            import_mode = self.import_mode_var.get()
            
            # 关闭导入选项对话框
            self.import_window.destroy()
            
            # 读取要导入的数据
            df_import = pd.read_excel(filepath, sheet_name=sheet_name)
            
            if import_mode == "merge":
                # 合并到当前工作表
                if self.df is not None:
                    # 检查列名是否匹配
                    if list(df_import.columns) == list(self.df.columns):
                        # 合并数据
                        self.df = pd.concat([self.df, df_import], ignore_index=True)
                        self.update_treeview()
                        self.status_var.set(f"已导入{len(df_import)}条记录，当前共{len(self.df)}条记录")
                        messagebox.showinfo("成功", f"成功导入{len(df_import)}条记录")
                    else:
                        messagebox.showerror("错误", "导入文件的列名与当前工作表不匹配")
                else:
                    messagebox.showwarning("警告", "请先打开一个Excel文件")
            else:
                # 创建新工作表（暂时只支持替换当前工作表，完整功能需要更复杂的ExcelWriter操作）
                self.df = df_import
                self.update_treeview()
                self.status_var.set(f"已创建新工作表，共{len(df_import)}条记录")
                messagebox.showinfo("成功", "成功创建新工作表")
                
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {str(e)}")
            self.status_var.set("导入失败")
    
    def export_file(self):
        """导出Excel文件"""
        if self.df is not None:
            # 打开文件保存对话框
            filetypes = [("Excel文件", "*.xlsx"), ("Excel 97-2003文件", "*.xls"), ("所有文件", "*.*")]
            filepath = filedialog.asksaveasfilename(title="导出Excel文件", filetypes=filetypes, defaultextension=".xlsx")
            
            if filepath:
                try:
                    self.status_var.set(f"正在导出文件: {filepath}")
                    self.root.update_idletasks()
                    
                    # 导出当前工作表数据
                    self.df.to_excel(filepath, sheet_name=self.current_sheet, index=False)
                    
                    self.status_var.set(f"文件已导出: {os.path.basename(filepath)}")
                    messagebox.showinfo("成功", "文件导出成功")
                except Exception as e:
                    messagebox.showerror("错误", f"导出文件失败: {str(e)}")
                    self.status_var.set("导出文件失败")
        else:
            messagebox.showwarning("警告", "没有可导出的数据")
        
    def add_record(self):
        """新增记录"""
        if self.df is not None:
            # 创建新增记录对话框
            self.add_window = tk.Toplevel(self.root)
            self.add_window.title("新增记录")
            self.add_window.geometry("400x300")
            self.add_window.resizable(False, False)
            
            # 居中显示
            self.add_window.transient(self.root)
            self.add_window.grab_set()
            
            # 创建滚动区域
            scroll_frame = ttk.Frame(self.add_window)
            scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            canvas = tk.Canvas(scroll_frame)
            scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # 创建输入框字典，用于存储各个字段的输入
            self.entry_vars = {}
            
            # 为每个字段创建标签和输入框
            for i, column in enumerate(self.df.columns):
                ttk.Label(scrollable_frame, text=column).grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
                var = tk.StringVar()
                entry = ttk.Entry(scrollable_frame, textvariable=var, width=30)
                entry.grid(row=i, column=1, padx=5, pady=5, sticky=tk.EW)
                self.entry_vars[column] = var
            
            # 创建按钮框架
            button_frame = ttk.Frame(self.add_window)
            button_frame.pack(fill=tk.X, padx=10, pady=10)
            
            # 创建确认和取消按钮
            ttk.Button(button_frame, text="确认", command=self.save_new_record).pack(side=tk.RIGHT, padx=5)
            ttk.Button(button_frame, text="取消", command=self.add_window.destroy).pack(side=tk.RIGHT, padx=5)
        else:
            messagebox.showwarning("警告", "请先打开一个Excel文件")
    
    def save_new_record(self):
        """保存新记录"""
        try:
            # 收集输入值并验证
            new_record = {}
            for column, var in self.entry_vars.items():
                value = var.get()
                
                # 数据验证
                if not self.validate_data(column, value):
                    messagebox.showerror("错误", f"数据类型不匹配！列 '{column}' 应输入{self.df[column].dtype}类型的数据")
                    return
                
                new_record[column] = value
            
            # 创建新的DataFrame行
            new_row = pd.DataFrame([new_record])
            
            # 将新行添加到原DataFrame
            self.df = pd.concat([self.df, new_row], ignore_index=True)
            
            # 更新Treeview
            self.update_treeview()
            
            # 关闭对话框
            self.add_window.destroy()
            
            self.status_var.set(f"已新增一条记录，共{len(self.df)}条记录")
        except Exception as e:
            messagebox.showerror("错误", f"新增记录失败: {str(e)}")
            self.status_var.set("新增记录失败")
        
    def delete_record(self):
        """删除记录"""
        if self.df is not None:
            # 获取选中的记录
            selected_items = self.tree.selection()
            
            if selected_items:
                # 弹出确认对话框
                confirm = messagebox.askyesno("确认删除", f"确定要删除选中的{len(selected_items)}条记录吗？")
                
                if confirm:
                    try:
                        # 获取选中记录的索引
                        selected_indices = [int(item_id) for item_id in selected_items]
                        
                        # 从DataFrame中删除记录
                        self.df = self.df.drop(selected_indices)
                        self.df = self.df.reset_index(drop=True)
                        
                        # 更新Treeview
                        self.update_treeview()
                        
                        self.status_var.set(f"已删除{len(selected_items)}条记录，剩余{len(self.df)}条记录")
                    except Exception as e:
                        messagebox.showerror("错误", f"删除记录失败: {str(e)}")
                        self.status_var.set("删除记录失败")
            else:
                messagebox.showwarning("警告", "请先选择要删除的记录")
        else:
            messagebox.showwarning("警告", "请先打开一个Excel文件")
        
    def modify_record(self):
        """修改记录"""
        if self.df is not None:
            # 获取选中的记录
            selected_items = self.tree.selection()
            
            if len(selected_items) == 1:
                item_id = selected_items[0]
                row_index = int(item_id)
                
                # 创建修改记录对话框
                self.modify_window = tk.Toplevel(self.root)
                self.modify_window.title("修改记录")
                self.modify_window.geometry("400x300")
                self.modify_window.resizable(False, False)
                
                # 居中显示
                self.modify_window.transient(self.root)
                self.modify_window.grab_set()
                
                # 创建滚动区域
                scroll_frame = ttk.Frame(self.modify_window)
                scroll_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                canvas = tk.Canvas(scroll_frame)
                scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
                scrollable_frame = ttk.Frame(canvas)
                
                scrollable_frame.bind(
                    "<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
                )
                
                canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")
                
                # 创建输入框字典，用于存储各个字段的输入
                self.entry_vars = {}
                
                # 获取当前记录的值
                current_row = self.df.iloc[row_index]
                
                # 为每个字段创建标签和输入框，并填充当前值
                for i, column in enumerate(self.df.columns):
                    ttk.Label(scrollable_frame, text=column).grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
                    var = tk.StringVar()
                    var.set(str(current_row[column]))
                    entry = ttk.Entry(scrollable_frame, textvariable=var, width=30)
                    entry.grid(row=i, column=1, padx=5, pady=5, sticky=tk.EW)
                    self.entry_vars[column] = var
                
                # 创建按钮框架
                button_frame = ttk.Frame(self.modify_window)
                button_frame.pack(fill=tk.X, padx=10, pady=10)
                
                # 创建确认和取消按钮，传递row_index参数
                ttk.Button(button_frame, text="确认", command=lambda: self.save_modified_record(row_index)).pack(side=tk.RIGHT, padx=5)
                ttk.Button(button_frame, text="取消", command=self.modify_window.destroy).pack(side=tk.RIGHT, padx=5)
            elif len(selected_items) > 1:
                messagebox.showwarning("警告", "一次只能修改一条记录")
            else:
                messagebox.showwarning("警告", "请先选择要修改的记录")
        else:
            messagebox.showwarning("警告", "请先打开一个Excel文件")
    
    def save_modified_record(self, row_index):
        """保存修改后的记录"""
        try:
            # 收集修改后的值并验证
            modified_record = {}
            for column, var in self.entry_vars.items():
                value = var.get()
                
                # 数据验证
                if not self.validate_data(column, value):
                    messagebox.showerror("错误", f"数据类型不匹配！列 '{column}' 应输入{self.df[column].dtype}类型的数据")
                    return
                
                modified_record[column] = value
            
            # 更新DataFrame中的记录
            for column, value in modified_record.items():
                self.df.at[row_index, column] = value
            
            # 更新Treeview
            self.update_treeview()
            
            # 关闭对话框
            self.modify_window.destroy()
            
            self.status_var.set(f"已修改第{row_index+1}行记录")
        except Exception as e:
            messagebox.showerror("错误", f"修改记录失败: {str(e)}")
            self.status_var.set("修改记录失败")
        
    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo("关于", "Excel文件管理器 v1.0\n\n用于管理本地Excel文件的GUI工具")
        
    def on_sheet_change(self, event):
        """工作表切换事件处理"""
        sheet_name = self.sheet_var.get()
        if sheet_name != self.current_sheet:
            self.current_sheet = sheet_name
            self.load_sheet_data(sheet_name)
        
    def load_sheet_data(self, sheet_name):
        """加载指定工作表的数据"""
        try:
            self.status_var.set(f"正在加载工作表: {sheet_name}")
            self.root.update_idletasks()
            
            # 读取工作表数据
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            
            # 更新Treeview显示
            self.update_treeview()
            
            self.status_var.set(f"已加载工作表: {sheet_name}，共{len(self.df)}条记录")
        except Exception as e:
            messagebox.showerror("错误", f"加载工作表失败: {str(e)}")
            self.status_var.set("加载工作表失败")
        
    def on_cell_double_click(self, event):
        """双击单元格事件处理，显示编辑框"""
        # 获取双击位置
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            # 获取选中的项和列
            item_id = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            
            if item_id and column:
                # 隐藏之前的编辑框
                self.edit_entry.pack_forget()
                
                # 获取列索引（从1开始）
                col_index = int(column.replace('#', '')) - 1
                if col_index < len(self.df.columns):
                    # 获取单元格值
                    current_value = self.tree.item(item_id, 'values')[col_index]
                    
                    # 获取单元格位置
                    x, y, width, height = self.tree.bbox(item_id, column)
                    
                    # 显示编辑框
                    self.edit_entry.place(x=x, y=y, width=width, height=height)
                    self.edit_entry.delete(0, tk.END)
                    self.edit_entry.insert(0, current_value)
                    self.edit_entry.focus_set()
                    self.edit_entry.select_range(0, tk.END)
                    
                    # 记录当前编辑的单元格
                    self.editing_cell = (item_id, self.df.columns[col_index])
        
    def on_edit_focus_out(self, event):
        """编辑框失去焦点事件处理，保存修改"""
        self.save_edit()
        
    def on_edit_return(self, event):
        """编辑框回车键事件处理，保存修改"""
        self.save_edit()
        
    def on_edit_escape(self, event):
        """编辑框ESC键事件处理，取消编辑"""
        self.edit_entry.pack_forget()
        self.editing_cell = None
        
    def validate_data(self, column, value):
        """验证数据类型是否匹配"""
        if pd.isna(value) or value == '':
            return True  # 允许空值
        
        try:
            # 获取原数据类型
            original_type = self.df[column].dtype
            
            # 尝试转换数据类型
            if original_type == 'int64' or original_type == 'float64':
                float(value)  # 尝试转换为数字
            elif original_type == 'datetime64[ns]':
                pd.to_datetime(value)  # 尝试转换为日期时间
            
            return True
        except:
            return False
    
    def save_edit(self):
        """保存编辑内容"""
        if self.editing_cell:
            item_id, column = self.editing_cell
            new_value = self.edit_entry.get()
            
            # 数据验证
            if not self.validate_data(column, new_value):
                messagebox.showerror("错误", f"数据类型不匹配！列 '{column}' 应输入{self.df[column].dtype}类型的数据")
                return
            
            # 更新Treeview显示
            col_index = list(self.df.columns).index(column)
            values = list(self.tree.item(item_id, 'values'))
            values[col_index] = new_value
            self.tree.item(item_id, values=values)
            
            # 更新DataFrame
            row_index = int(item_id)
            self.df.at[row_index, column] = new_value
            
            # 隐藏编辑框
            self.edit_entry.pack_forget()
            self.editing_cell = None
            
            self.status_var.set(f"已修改单元格: 行{row_index+1}, 列{column}")
        
    def update_treeview(self):
        """更新Treeview数据"""
        # 清空Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 清空列
        self.tree['columns'] = []
        for col in self.tree['columns']:
            self.tree.heading(col, text='')
            self.tree.column(col, width=0)
        
        if self.df is not None and not self.df.empty:
            # 设置列
            self.tree['columns'] = list(self.df.columns)
            
            # 设置列标题和排序功能
            for col in self.df.columns:
                self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(_col, False))
                
                # 自动调整列宽
                max_width = 100  # 最小宽度
                # 检查列名长度
                if len(col) > max_width // 8:  # 假设每个字符宽度为8像素
                    max_width = len(col) * 8 + 20  # 20像素边距
                # 检查数据宽度
                for index, row in self.df.iterrows():
                    cell_value = str(row[col])
                    if len(cell_value) > max_width // 8:
                        max_width = len(cell_value) * 8 + 20
                
                self.tree.column(col, width=min(max_width, 300), anchor=tk.CENTER, stretch=True)
            
            # 添加数据行
            for index, row in self.df.iterrows():
                values = [str(row[col]) for col in self.df.columns]
                self.tree.insert('', tk.END, values=values, iid=str(index))
            
            # 更新状态栏显示
            self.status_var.set(f"已加载工作表: {self.current_sheet}，共{len(self.df)}条记录")
    
    def sort_column(self, column, reverse):
        """排序列"""
        if self.df is not None and not self.df.empty:
            # 对DataFrame进行排序
            self.df = self.df.sort_values(by=column, ascending=not reverse)
            self.df = self.df.reset_index(drop=True)
            
            # 更新Treeview
            self.update_treeview()
            
            # 更新列标题的排序指示
            for col in self.df.columns:
                if col == column:
                    self.tree.heading(col, text=f"{col} {'↑' if reverse else '↓'}", 
                                     command=lambda _col=col: self.sort_column(_col, not reverse))
                else:
                    self.tree.heading(col, text=col, 
                                     command=lambda _col=col: self.sort_column(_col, False))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelManager(root)
    root.mainloop()
