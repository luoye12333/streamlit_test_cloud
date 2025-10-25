import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import win32file
import win32con
import pywintypes
from pathlib import Path
import threading
import time

class BatchFileTimeModifier:
    def __init__(self, root):
        self.root = root
        self.root.title("批量修改文件时间戳工具")
        self.root.geometry("900x700")
        
        self.folder_path = None
        self.is_processing = False
        self.total_files = 0
        self.processed_files = 0
        self.failed_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # 文件夹选择区域
        folder_frame = ttk.LabelFrame(main_frame, text="文件夹选择", padding="5")
        folder_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        self.folder_path_var = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.folder_path_var, width=80, state="readonly").grid(row=0, column=0, padx=(0, 5))
        ttk.Button(folder_frame, text="选择文件夹", command=self.select_folder).grid(row=0, column=1)
        ttk.Button(folder_frame, text="扫描文件", command=self.scan_files).grid(row=0, column=2, padx=(5, 0))
        
        # 选项设置区域
        options_frame = ttk.LabelFrame(main_frame, text="处理选项", padding="5")
        options_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        # 文件过滤选项
        filter_frame = ttk.Frame(options_frame)
        filter_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        ttk.Label(filter_frame, text="文件类型过滤:").grid(row=0, column=0, sticky=tk.W)
        self.file_filter_var = tk.StringVar(value="*.*")
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.file_filter_var, width=15)
        filter_combo['values'] = ("*.*", "*.txt", "*.doc", "*.docx", "*.pdf", "*.jpg", "*.png", "*.mp4", "*.mp3")
        filter_combo.grid(row=0, column=1, padx=(5, 20), sticky=tk.W)
        
        # 递归处理选项
        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(filter_frame, text="包含子文件夹", variable=self.recursive_var).grid(row=0, column=2, sticky=tk.W)
        
        # 隐藏文件选项
        self.include_hidden_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(filter_frame, text="包含隐藏文件", variable=self.include_hidden_var).grid(row=0, column=3, padx=(20, 0), sticky=tk.W)
        
        # 时间设置区域
        time_frame = ttk.LabelFrame(main_frame, text="时间设置", padding="5")
        time_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        # 创建时间设置
        ttk.Label(time_frame, text="创建时间:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.create_time_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        create_time_entry = ttk.Entry(time_frame, textvariable=self.create_time_var, width=25)
        create_time_entry.grid(row=0, column=1, padx=(5, 10), pady=2)
        ttk.Button(time_frame, text="当前时间", command=lambda: self.set_current_time("create")).grid(row=0, column=2, pady=2)
        
        self.modify_create_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(time_frame, text="修改创建时间", variable=self.modify_create_var).grid(row=0, column=3, padx=(10, 0), pady=2)
        
        # 修改时间设置
        ttk.Label(time_frame, text="修改时间:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.modify_time_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        modify_time_entry = ttk.Entry(time_frame, textvariable=self.modify_time_var, width=25)
        modify_time_entry.grid(row=1, column=1, padx=(5, 10), pady=2)
        ttk.Button(time_frame, text="当前时间", command=lambda: self.set_current_time("modify")).grid(row=1, column=2, pady=2)
        
        self.modify_modify_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(time_frame, text="修改修改时间", variable=self.modify_modify_var).grid(row=1, column=3, padx=(10, 0), pady=2)
        
        # 访问时间设置
        ttk.Label(time_frame, text="访问时间:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.access_time_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        access_time_entry = ttk.Entry(time_frame, textvariable=self.access_time_var, width=25)
        access_time_entry.grid(row=2, column=1, padx=(5, 10), pady=2)
        ttk.Button(time_frame, text="当前时间", command=lambda: self.set_current_time("access")).grid(row=2, column=2, pady=2)
        
        self.modify_access_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(time_frame, text="修改访问时间", variable=self.modify_access_var).grid(row=2, column=3, padx=(10, 0), pady=2)
        
        # 文件列表显示区域
        list_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="5")
        list_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        
        # 创建文件列表树形视图
        columns = ("文件名", "路径", "大小", "创建时间", "修改时间", "状态")
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)
        
        # 设置列标题和宽度
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("路径", text="路径")
        self.file_tree.heading("大小", text="大小")
        self.file_tree.heading("创建时间", text="创建时间")
        self.file_tree.heading("修改时间", text="修改时间")
        self.file_tree.heading("状态", text="状态")
        
        self.file_tree.column("文件名", width=150)
        self.file_tree.column("路径", width=200)
        self.file_tree.column("大小", width=80)
        self.file_tree.column("创建时间", width=120)
        self.file_tree.column("修改时间", width=120)
        self.file_tree.column("状态", width=80)
        
        # 添加滚动条
        tree_scrollbar_v = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_tree.yview)
        tree_scrollbar_h = ttk.Scrollbar(list_frame, orient="horizontal", command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=tree_scrollbar_v.set, xscrollcommand=tree_scrollbar_h.set)
        
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        tree_scrollbar_v.grid(row=0, column=1, sticky="ns")
        tree_scrollbar_h.grid(row=1, column=0, sticky="ew")
        
        # 统计信息
        stats_frame = ttk.Frame(main_frame)
        stats_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        self.stats_label = ttk.Label(stats_frame, text="请选择文件夹并扫描文件")
        self.stats_label.grid(row=0, column=0, sticky=tk.W)
        
        # 进度条
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        ttk.Label(progress_frame, text="处理进度:").grid(row=0, column=0, sticky=tk.W)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, length=400)
        self.progress_bar.grid(row=0, column=1, padx=(5, 5), sticky="ew")
        
        self.progress_label = ttk.Label(progress_frame, text="0/0")
        self.progress_label.grid(row=0, column=2, sticky=tk.W)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        self.start_button = ttk.Button(button_frame, text="开始批量修改", command=self.start_batch_modify)
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="停止处理", command=self.stop_processing, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(button_frame, text="清空列表", command=self.clear_file_list).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(button_frame, text="导出日志", command=self.export_log).grid(row=0, column=3, padx=(0, 10))
        ttk.Button(button_frame, text="关于", command=self.show_about).grid(row=0, column=4)
        
        # 配置网格权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        progress_frame.columnconfigure(1, weight=1)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def select_folder(self):
        """选择文件夹"""
        folder_path = filedialog.askdirectory(title="选择要批量修改文件时间的文件夹")
        
        if folder_path:
            self.folder_path = folder_path
            self.folder_path_var.set(folder_path)
            self.clear_file_list()
    
    def scan_files(self):
        """扫描文件夹中的文件"""
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择一个文件夹")
            return
        
        if not os.path.exists(self.folder_path):
            messagebox.showerror("错误", "文件夹不存在")
            return
        
        # 清空现有列表
        self.clear_file_list()
        
        try:
            # 获取文件过滤器
            file_filter = self.file_filter_var.get().strip()
            if not file_filter:
                file_filter = "*.*"
            
            # 扫描文件
            files_found = []
            folder_path = Path(self.folder_path)
            
            if self.recursive_var.get():
                # 递归扫描
                if file_filter == "*.*":
                    pattern = "**/*"
                else:
                    pattern = f"**/{file_filter}"
                files_found = list(folder_path.glob(pattern))
            else:
                # 只扫描当前文件夹
                if file_filter == "*.*":
                    pattern = "*"
                else:
                    pattern = file_filter
                files_found = list(folder_path.glob(pattern))
            
            # 过滤文件（排除文件夹，处理隐藏文件选项）
            valid_files = []
            for file_path in files_found:
                if file_path.is_file():
                    # 检查是否为隐藏文件
                    if not self.include_hidden_var.get():
                        if file_path.name.startswith('.') or self.is_hidden_file(str(file_path)):
                            continue
                    valid_files.append(file_path)
            
            # 添加到列表中
            for file_path in valid_files:
                try:
                    stat_info = file_path.stat()
                    create_time = datetime.fromtimestamp(stat_info.st_ctime).strftime("%Y-%m-%d %H:%M:%S")
                    modify_time = datetime.fromtimestamp(stat_info.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    
                    self.file_tree.insert("", "end", values=(
                        file_path.name,
                        str(file_path.parent),
                        self.format_size(stat_info.st_size),
                        create_time,
                        modify_time,
                        "待处理"
                    ))
                except Exception as e:
                    self.file_tree.insert("", "end", values=(
                        file_path.name,
                        str(file_path.parent),
                        "错误",
                        "错误",
                        "错误",
                        f"读取失败: {str(e)}"
                    ))
            
            # 更新统计信息
            total_files = len(valid_files)
            self.total_files = total_files
            self.stats_label.config(text=f"共找到 {total_files} 个文件")
            
            if total_files == 0:
                messagebox.showinfo("信息", "在指定文件夹中没有找到符合条件的文件")
            
        except Exception as e:
            messagebox.showerror("错误", f"扫描文件夹失败: {str(e)}")
    
    def is_hidden_file(self, file_path):
        """检查文件是否为隐藏文件（Windows）"""
        try:
            import win32api
            import win32con
            attrs = win32api.GetFileAttributes(file_path)
            return bool(attrs & win32con.FILE_ATTRIBUTE_HIDDEN)
        except:
            return False
    
    def format_size(self, size_bytes):
        """格式化文件大小"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.2f} {size_names[i]}"
    
    def set_current_time(self, time_type):
        """设置当前时间"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if time_type == "create":
            self.create_time_var.set(current_time)
        elif time_type == "modify":
            self.modify_time_var.set(current_time)
        elif time_type == "access":
            self.access_time_var.set(current_time)
    
    def clear_file_list(self):
        """清空文件列表"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        self.total_files = 0
        self.processed_files = 0
        self.failed_files = []
        self.stats_label.config(text="请选择文件夹并扫描文件")
        self.progress_var.set(0)
        self.progress_label.config(text="0/0")
    
    def start_batch_modify(self):
        """开始批量修改"""
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        
        if self.total_files == 0:
            messagebox.showwarning("警告", "没有找到要处理的文件，请先扫描文件")
            return
        
        # 验证时间格式
        try:
            if self.modify_create_var.get():
                datetime.strptime(self.create_time_var.get(), "%Y-%m-%d %H:%M:%S")
            if self.modify_modify_var.get():
                datetime.strptime(self.modify_time_var.get(), "%Y-%m-%d %H:%M:%S")
            if self.modify_access_var.get():
                datetime.strptime(self.access_time_var.get(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            messagebox.showerror("错误", "时间格式错误，请使用格式: YYYY-MM-DD HH:MM:SS")
            return
        
        # 检查是否至少选择了一个时间类型进行修改
        if not (self.modify_create_var.get() or self.modify_modify_var.get() or self.modify_access_var.get()):
            messagebox.showwarning("警告", "请至少选择一个时间类型进行修改")
            return
        
        # 确认对话框
        message = f"确定要修改 {self.total_files} 个文件的时间戳吗？\n\n"
        if self.modify_create_var.get():
            message += f"创建时间: {self.create_time_var.get()}\n"
        if self.modify_modify_var.get():
            message += f"修改时间: {self.modify_time_var.get()}\n"
        if self.modify_access_var.get():
            message += f"访问时间: {self.access_time_var.get()}\n"
        
        if not messagebox.askyesno("确认", message):
            return
        
        # 开始处理
        self.is_processing = True
        self.processed_files = 0
        self.failed_files = []
        
        # 禁用开始按钮，启用停止按钮
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        # 在新线程中处理文件
        thread = threading.Thread(target=self.process_files_thread)
        thread.daemon = True
        thread.start()
    
    def process_files_thread(self):
        """在后台线程中处理文件"""
        try:
            # 获取所有文件信息
            files_to_process = []
            for item in self.file_tree.get_children():
                values = self.file_tree.item(item, "values")
                file_name = values[0]
                file_dir = values[1]
                file_path = os.path.join(file_dir, file_name)
                files_to_process.append((item, file_path))
            
            # 解析目标时间
            target_times = {}
            if self.modify_create_var.get():
                target_times['create'] = datetime.strptime(self.create_time_var.get(), "%Y-%m-%d %H:%M:%S")
            if self.modify_modify_var.get():
                target_times['modify'] = datetime.strptime(self.modify_time_var.get(), "%Y-%m-%d %H:%M:%S")
            if self.modify_access_var.get():
                target_times['access'] = datetime.strptime(self.access_time_var.get(), "%Y-%m-%d %H:%M:%S")
            
            # 处理每个文件
            for item, file_path in files_to_process:
                if not self.is_processing:  # 检查是否被停止
                    break
                
                try:
                    # 更新界面显示当前处理的文件
                    self.root.after(0, lambda i=item: self.file_tree.set(i, "状态", "处理中..."))
                    
                    # 修改文件时间
                    success = self.modify_file_time(file_path, target_times)
                    
                    if success:
                        self.processed_files += 1
                        self.root.after(0, lambda i=item: self.file_tree.set(i, "状态", "成功"))
                        
                        # 更新显示的时间
                        if 'create' in target_times:
                            self.root.after(0, lambda i=item, t=target_times['create'].strftime("%Y-%m-%d %H:%M:%S"): 
                                          self.file_tree.set(i, "创建时间", t))
                        if 'modify' in target_times:
                            self.root.after(0, lambda i=item, t=target_times['modify'].strftime("%Y-%m-%d %H:%M:%S"): 
                                          self.file_tree.set(i, "修改时间", t))
                    else:
                        self.failed_files.append(file_path)
                        self.root.after(0, lambda i=item: self.file_tree.set(i, "状态", "失败"))
                    
                except Exception as e:
                    self.failed_files.append(f"{file_path}: {str(e)}")
                    self.root.after(0, lambda i=item, e=str(e): self.file_tree.set(i, "状态", f"错误: {e}"))
                
                # 更新进度
                progress = (self.processed_files + len(self.failed_files)) / self.total_files * 100
                self.root.after(0, lambda p=progress: self.progress_var.set(p))
                self.root.after(0, lambda: self.progress_label.config(
                    text=f"{self.processed_files + len(self.failed_files)}/{self.total_files}"))
                
                # 短暂延迟，避免过快处理
                time.sleep(0.01)
            
            # 处理完成
            self.root.after(0, self.processing_completed)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"批量处理过程中发生错误: {str(e)}"))
            self.root.after(0, self.processing_completed)
    
    def modify_file_time(self, file_path, target_times):
        """修改单个文件的时间"""
        try:
            # 首先使用os.utime修改访问和修改时间
            if 'access' in target_times or 'modify' in target_times:
                current_stat = os.stat(file_path)
                access_time = target_times.get('access', datetime.fromtimestamp(current_stat.st_atime))
                modify_time = target_times.get('modify', datetime.fromtimestamp(current_stat.st_mtime))
                
                os.utime(file_path, (access_time.timestamp(), modify_time.timestamp()))
            
            # 如果需要修改创建时间，使用Windows API
            if 'create' in target_times:
                handle = win32file.CreateFile(
                    file_path,
                    win32con.GENERIC_WRITE,
                    win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                    None,
                    win32con.OPEN_EXISTING,
                    0,
                    None
                )
                
                try:
                    # 转换为win32时间格式
                    create_time_win32 = pywintypes.Time(target_times['create'].timestamp())
                    
                    # 获取当前的访问和修改时间（如果不需要修改的话）
                    if 'access' not in target_times or 'modify' not in target_times:
                        current_stat = os.stat(file_path)
                        if 'access' not in target_times:
                            access_time_win32 = pywintypes.Time(current_stat.st_atime)
                        else:
                            access_time_win32 = pywintypes.Time(target_times['access'].timestamp())
                        
                        if 'modify' not in target_times:
                            modify_time_win32 = pywintypes.Time(current_stat.st_mtime)
                        else:
                            modify_time_win32 = pywintypes.Time(target_times['modify'].timestamp())
                    else:
                        access_time_win32 = pywintypes.Time(target_times['access'].timestamp())
                        modify_time_win32 = pywintypes.Time(target_times['modify'].timestamp())
                    
                    # 设置文件时间
                    win32file.SetFileTime(handle.handle, create_time_win32, access_time_win32, modify_time_win32)
                    
                finally:
                    handle.Close()
            
            return True
            
        except Exception as e:
            print(f"修改文件 {file_path} 时间失败: {str(e)}")
            return False
    
    def stop_processing(self):
        """停止处理"""
        self.is_processing = False
        self.processing_completed()
    
    def processing_completed(self):
        """处理完成"""
        self.is_processing = False
        
        # 恢复按钮状态
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        # 显示结果
        success_count = self.processed_files
        failed_count = len(self.failed_files)
        total_count = success_count + failed_count
        
        message = f"批量处理完成！\n\n"
        message += f"总计: {total_count} 个文件\n"
        message += f"成功: {success_count} 个文件\n"
        message += f"失败: {failed_count} 个文件"
        
        if failed_count > 0:
            message += f"\n\n失败的文件:\n"
            for i, failed_file in enumerate(self.failed_files[:5]):  # 只显示前5个
                message += f"{i+1}. {failed_file}\n"
            if failed_count > 5:
                message += f"... 还有 {failed_count - 5} 个文件失败"
        
        messagebox.showinfo("处理完成", message)
    
    def export_log(self):
        """导出处理日志"""
        if self.total_files == 0:
            messagebox.showwarning("警告", "没有处理记录可导出")
            return
        
        try:
            log_file = filedialog.asksaveasfilename(
                title="保存处理日志",
                defaultextension=".txt",
                filetypes=[("\u6587\u672c\u6587\u4ef6", "*.txt"), ("\u6240\u6709\u6587\u4ef6", "*.*")]
            )
            
            if log_file:
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write("批量文件时间修改日志\n")
                    f.write("=" * 50 + "\n")
                    f.write(f"处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"处理文件夹: {self.folder_path}\n")
                    f.write(f"文件过滤器: {self.file_filter_var.get()}\n")
                    f.write(f"包含子文件夹: {'是' if self.recursive_var.get() else '否'}\n")
                    f.write(f"包含隐藏文件: {'是' if self.include_hidden_var.get() else '否'}\n")
                    f.write("\n时间设置:\n")
                    if self.modify_create_var.get():
                        f.write(f"创建时间: {self.create_time_var.get()}\n")
                    if self.modify_modify_var.get():
                        f.write(f"修改时间: {self.modify_time_var.get()}\n")
                    if self.modify_access_var.get():
                        f.write(f"访问时间: {self.access_time_var.get()}\n")
                    
                    f.write("\n处理结果:\n")
                    f.write(f"成功: {self.processed_files} 个文件\n")
                    f.write(f"失败: {len(self.failed_files)} 个文件\n")
                    
                    if self.failed_files:
                        f.write("\n失败文件列表:\n")
                        for i, failed_file in enumerate(self.failed_files, 1):
                            f.write(f"{i}. {failed_file}\n")
                    
                    f.write("\n所有文件列表:\n")
                    for item in self.file_tree.get_children():
                        values = self.file_tree.item(item, "values")
                        f.write(f"{values[0]} | {values[1]} | {values[5]}\n")
                
                messagebox.showinfo("成功", f"日志已导出到: {log_file}")
        
        except Exception as e:
            messagebox.showerror("错误", f"导出日志失败: {str(e)}")
    
    def show_about(self):
        """显示关于信息"""
        about_text = """批量文件时间戳修改工具 v1.0

功能特色：
\u2022 支持批量修改文件的创建时间、修改时间和访问时间
\u2022 支持文件类型过滤和递归扫描
\u2022 可选择包含或排除隐藏文件
\u2022 实时显示处理进度和结果
\u2022 支持导出详细处理日志
\u2022 支持停止正在进行的处理

注意事项：
\u2022 修改文件时间可能需要管理员权限
\u2022 建议在修改前备份重要文件
\u2022 时间格式必须为: YYYY-MM-DD HH:MM:SS

作者： AI Assistant
版本： 1.0
日期： 2025-01-22"""
        
        messagebox.showinfo("关于", about_text)

def main():
    """主程序入口"""
    root = tk.Tk()
    app = BatchFileTimeModifier(root)
    root.mainloop()

if __name__ == "__main__":
    main()