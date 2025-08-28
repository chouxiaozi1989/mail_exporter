import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import threading
import queue
import os
from mail_exporter import fetch_emails, get_mail_folders

# 尝试导入日期选择器，如果没有安装则使用普通输入框
try:
    from tkcalendar import DateEntry
    HAS_DATE_PICKER = True
except ImportError:
    HAS_DATE_PICKER = False

# 程序版本信息
VERSION = "1.2.0"

class MailExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("163邮箱邮件导出工具")
        self.root.geometry("700x650")
        self.root.resizable(True, True)
        
        # 创建消息队列用于线程间通信
        self.message_queue = queue.Queue()
        self.export_thread = None
        self.is_exporting = False
        self.folder_mapping = {}  # 文件夹显示名到实际名的映射
        
        self.setup_ui()
        self.check_queue()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题和版本信息
        title_label = ttk.Label(main_frame, text="163邮箱邮件导出工具", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 5))
        
        version_label = ttk.Label(main_frame, text=f"版本 {VERSION}", font=('Arial', 9), foreground='gray')
        version_label.grid(row=1, column=0, columnspan=2, pady=(0, 15))
        
        # 创建输入字段区域
        self.create_input_fields(main_frame)
        
        # 创建进度显示区域
        self.create_progress_area(main_frame)
        
        # 创建操作按钮区域
        self.create_action_buttons(main_frame)
        
        # 创建结果显示区域
        self.create_result_area(main_frame)
    
    def create_input_fields(self, parent):
        """创建输入字段"""
        # 输入字段框架
        input_frame = ttk.LabelFrame(parent, text="邮箱配置", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # 用户名
        ttk.Label(input_frame, text="邮箱用户名:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(input_frame, textvariable=self.username_var, width=30)
        username_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        row += 1
        
        # 密码
        ttk.Label(input_frame, text="密码/授权码:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(input_frame, textvariable=self.password_var, show="*", width=30)
        password_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        row += 1
        
        # 开始日期
        ttk.Label(input_frame, text="开始日期:").grid(row=row, column=0, sticky=tk.W, pady=2)
        date_frame = ttk.Frame(input_frame)
        date_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        today = datetime.now()
        first_day = datetime(today.year, today.month, 1)
        
        if HAS_DATE_PICKER:
            # 使用日期选择器
            self.start_date_picker = DateEntry(date_frame, width=12, background='darkblue',
                                             foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.start_date_picker.set_date(first_day)
            self.start_date_picker.grid(row=0, column=0)
            
            ttk.Label(date_frame, text="至").grid(row=0, column=1, padx=5)
            
            self.end_date_picker = DateEntry(date_frame, width=12, background='darkblue',
                                           foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            self.end_date_picker.set_date(today)
            self.end_date_picker.grid(row=0, column=2)
        else:
            # 使用普通输入框
            self.start_date_var = tk.StringVar()
            self.start_date_var.set(first_day.strftime("%Y-%m-%d"))
            start_date_entry = ttk.Entry(date_frame, textvariable=self.start_date_var, width=12)
            start_date_entry.grid(row=0, column=0)
            
            ttk.Label(date_frame, text="至").grid(row=0, column=1, padx=5)
            
            self.end_date_var = tk.StringVar()
            self.end_date_var.set(today.strftime("%Y-%m-%d"))
            end_date_entry = ttk.Entry(date_frame, textvariable=self.end_date_var, width=12)
            end_date_entry.grid(row=0, column=2)
            
            ttk.Label(date_frame, text="(格式: YYYY-MM-DD)").grid(row=0, column=3, padx=(10, 0))
        
        row += 1
        
        # 邮箱文件夹
        ttk.Label(input_frame, text="邮箱文件夹:").grid(row=row, column=0, sticky=tk.W, pady=2)
        folder_frame = ttk.Frame(input_frame)
        folder_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        folder_frame.columnconfigure(0, weight=1)
        
        self.folder_var = tk.StringVar(value="INBOX")
        self.folder_combobox = ttk.Combobox(folder_frame, textvariable=self.folder_var, state="readonly", width=25)
        self.folder_combobox['values'] = ["INBOX"]
        self.folder_combobox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        refresh_folder_btn = ttk.Button(folder_frame, text="刷新", command=self.refresh_folders)
        refresh_folder_btn.grid(row=0, column=1, padx=(5, 0))
        row += 1
        
        # 输出文件
        ttk.Label(input_frame, text="输出文件:").grid(row=row, column=0, sticky=tk.W, pady=2)
        output_frame = ttk.Frame(input_frame)
        output_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        output_frame.columnconfigure(0, weight=1)
        
        self.output_var = tk.StringVar(value="emails.csv")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var)
        output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        browse_btn = ttk.Button(output_frame, text="浏览", command=self.browse_output_file)
        browse_btn.grid(row=0, column=1, padx=(5, 0))
        row += 1
        
        # 附件下载选项
        ttk.Label(input_frame, text="附件下载:").grid(row=row, column=0, sticky=tk.W, pady=2)
        attachment_frame = ttk.Frame(input_frame)
        attachment_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        attachment_frame.columnconfigure(0, weight=1)
        
        self.download_attachments_var = tk.BooleanVar(value=False)
        attachment_checkbox = ttk.Checkbutton(attachment_frame, text="下载邮件附件", 
                                            variable=self.download_attachments_var,
                                            command=self.toggle_attachment_folder)
        attachment_checkbox.grid(row=0, column=0, sticky=tk.W)
        row += 1
        
        # 附件保存目录
        ttk.Label(input_frame, text="附件目录:").grid(row=row, column=0, sticky=tk.W, pady=2)
        attachment_folder_frame = ttk.Frame(input_frame)
        attachment_folder_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        attachment_folder_frame.columnconfigure(0, weight=1)
        
        self.attachment_folder_var = tk.StringVar(value="")
        self.attachment_folder_entry = ttk.Entry(attachment_folder_frame, textvariable=self.attachment_folder_var, state="disabled")
        self.attachment_folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.browse_attachment_btn = ttk.Button(attachment_folder_frame, text="浏览", command=self.browse_attachment_folder, state="disabled")
        self.browse_attachment_btn.grid(row=0, column=1, padx=(5, 0))
        
        # 添加说明文本
        attachment_note = ttk.Label(input_frame, text="(留空则在CSV文件同目录下创建attachments文件夹)", 
                                  font=('Arial', 8), foreground='gray')
        attachment_note.grid(row=row+1, column=1, sticky=tk.W, padx=(10, 0), pady=(0, 5))
    
    def create_progress_area(self, parent):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(parent, text="导出进度", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 状态文本
        self.status_var = tk.StringVar(value="准备就绪")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.grid(row=1, column=0, sticky=tk.W)
    
    def create_action_buttons(self, parent):
        """创建操作按钮"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(0, 10))
        
        self.start_btn = ttk.Button(button_frame, text="开始导出", command=self.start_export)
        self.start_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.cancel_btn = ttk.Button(button_frame, text="取消", command=self.cancel_export, state="disabled")
        self.cancel_btn.grid(row=0, column=1)
    
    def create_result_area(self, parent):
        """创建结果显示区域"""
        result_frame = ttk.LabelFrame(parent, text="导出日志", padding="10")
        result_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 文本框和滚动条
        text_frame = ttk.Frame(result_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.result_text = tk.Text(text_frame, height=15, wrap=tk.WORD)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.result_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # 配置主框架的行权重
        parent.rowconfigure(5, weight=1)
    
    def refresh_folders(self):
        """刷新邮箱文件夹列表"""
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        
        if not username or not password:
            messagebox.showwarning("提示", "请先输入邮箱用户名和密码")
            return
        
        # 显示加载状态
        original_text = self.folder_combobox['state']
        self.folder_combobox.config(state="disabled")
        self.status_var.set("正在获取文件夹列表...")
        
        def get_folders_worker():
            try:
                folders = get_mail_folders(username, password)
                
                # 更新下拉列表
                folder_values = []
                folder_mapping = {}
                
                for folder_name, display_name in folders:
                    folder_values.append(display_name)
                    folder_mapping[display_name] = folder_name
                
                # 在主线程中更新UI
                self.root.after(0, lambda: self.update_folder_list(folder_values, folder_mapping))
                
            except Exception as e:
                error_msg = f"获取文件夹列表失败: {str(e)}"
                self.root.after(0, lambda: self.show_folder_error(error_msg))
        
        # 在后台线程中获取文件夹
        threading.Thread(target=get_folders_worker, daemon=True).start()
    
    def update_folder_list(self, folder_values, folder_mapping):
        """更新文件夹列表"""
        self.folder_combobox['values'] = folder_values
        self.folder_mapping = folder_mapping
        
        # 如果当前选择的文件夹不在新列表中，选择第一个
        current_value = self.folder_var.get()
        if current_value not in folder_values and folder_values:
            self.folder_var.set(folder_values[0])
        
        self.folder_combobox.config(state="readonly")
        self.status_var.set(f"已获取 {len(folder_values)} 个文件夹")
        self.log_message(f"成功获取 {len(folder_values)} 个邮箱文件夹")
    
    def show_folder_error(self, error_msg):
        """显示文件夹获取错误"""
        self.folder_combobox.config(state="readonly")
        self.status_var.set("准备就绪")
        messagebox.showerror("错误", error_msg)
        self.log_message(error_msg)
    
    def browse_output_file(self):
        """浏览输出文件"""
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if filename:
            self.output_var.set(filename)
    
    def toggle_attachment_folder(self):
        """切换附件文件夹输入框状态"""
        if self.download_attachments_var.get():
            self.attachment_folder_entry.config(state="normal")
            self.browse_attachment_btn.config(state="normal")
        else:
            self.attachment_folder_entry.config(state="disabled")
            self.browse_attachment_btn.config(state="disabled")
    
    def browse_attachment_folder(self):
        """浏览附件保存目录"""
        folder = filedialog.askdirectory(
            title="选择附件保存目录"
        )
        if folder:
            self.attachment_folder_var.set(folder)
    
    def start_export(self):
        """开始导出"""
        # 验证输入
        if not self.validate_inputs():
            return
        
        # 禁用开始按钮，启用取消按钮
        self.start_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.is_exporting = True
        
        # 清空结果文本
        self.result_text.delete(1.0, tk.END)
        
        # 重置进度
        self.progress_var.set(0)
        self.status_var.set("正在连接邮箱服务器...")
        
        # 启动导出线程
        self.export_thread = threading.Thread(target=self.export_worker, daemon=True)
        self.export_thread.start()
    
    def cancel_export(self):
        """取消导出"""
        self.is_exporting = False
        self.start_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.status_var.set("已取消")
        self.log_message("导出已取消")
    
    def validate_inputs(self):
        """验证输入"""
        if not self.username_var.get().strip():
            messagebox.showerror("错误", "请输入邮箱用户名")
            return False
        
        if not self.password_var.get().strip():
            messagebox.showerror("错误", "请输入密码或授权码")
            return False
        
        try:
            if HAS_DATE_PICKER:
                # 日期选择器已经确保了正确的日期格式
                start_date = self.start_date_picker.get_date()
                end_date = self.end_date_picker.get_date()
            else:
                # 验证手动输入的日期格式
                datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
                datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用YYYY-MM-DD格式")
            return False
        
        if not self.output_var.get().strip():
            messagebox.showerror("错误", "请指定输出文件")
            return False
        
        return True
    
    def progress_callback(self, current, total, message):
        """进度回调函数"""
        if not self.is_exporting:
            return
        
        # 计算进度百分比
        if total > 0:
            progress = min(100, max(0, (current / total) * 100))  # 确保进度在0-100之间
            self.message_queue.put(("progress", progress))
        elif current == 0 and total == 0:
            # 初始状态或准备阶段
            self.message_queue.put(("progress", 0))
        
        # 更新状态消息
        self.message_queue.put(("status", message))
        
        # 格式化日志消息
        if total > 0:
            self.message_queue.put(("log", f"[{current}/{total}] {message}"))
        else:
            self.message_queue.put(("log", f"[准备中] {message}"))
    
    def export_worker(self):
        """导出工作线程"""
        try:
            # 获取参数
            username = self.username_var.get().strip()
            password = self.password_var.get().strip()
            
            if HAS_DATE_PICKER:
                start_date = datetime.combine(self.start_date_picker.get_date(), datetime.min.time())
                end_date = datetime.combine(self.end_date_picker.get_date(), datetime.min.time())
            else:
                start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
                end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
            
            output_file = self.output_var.get().strip()
            
            # 获取实际的文件夹名（从显示名映射到实际名）
            folder_display = self.folder_var.get().strip() or "INBOX"
            folder = self.folder_mapping.get(folder_display, folder_display)
            
            # 如果没有映射，直接使用显示名作为文件夹名
            if not folder:
                folder = "INBOX"
            
            # 获取附件下载参数
            download_attachments = self.download_attachments_var.get()
            attachment_folder = self.attachment_folder_var.get().strip() if download_attachments else None
            
            # 发送开始消息
            self.message_queue.put(("log", f"开始导出邮件..."))
            self.message_queue.put(("log", f"用户: {username}"))
            self.message_queue.put(("log", f"时间范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}"))
            self.message_queue.put(("log", f"输出文件: {output_file}"))
            self.message_queue.put(("log", f"邮箱文件夹: {folder}"))
            
            if download_attachments:
                if attachment_folder:
                    self.message_queue.put(("log", f"附件保存目录: {attachment_folder}"))
                else:
                    self.message_queue.put(("log", "附件保存目录: CSV文件同目录下的attachments文件夹"))
            else:
                self.message_queue.put(("log", "不下载附件"))
            
            # 调用导出函数，传入进度回调和附件参数
            email_count = fetch_emails(username, password, start_date, end_date, output_file, folder, 
                                     self.progress_callback, download_attachments, attachment_folder)
            
            if self.is_exporting:
                self.message_queue.put(("success", f"导出完成! 共导出 {email_count} 封邮件到 {output_file}"))
                self.message_queue.put(("progress", 100))
                self.message_queue.put(("status", "导出完成"))
            
        except Exception as e:
            if self.is_exporting:
                self.message_queue.put(("error", f"导出失败: {str(e)}"))
                self.message_queue.put(("status", "导出失败"))
        finally:
            if self.is_exporting:
                self.message_queue.put(("finished", None))
    
    def check_queue(self):
        """检查消息队列"""
        try:
            while True:
                msg_type, msg_data = self.message_queue.get_nowait()
                
                if msg_type == "log":
                    self.log_message(msg_data)
                elif msg_type == "progress":
                    # 确保进度值在有效范围内
                    progress_value = max(0, min(100, msg_data))
                    self.progress_var.set(progress_value)
                elif msg_type == "status":
                    self.status_var.set(msg_data)
                elif msg_type == "success":
                    self.log_message(msg_data)
                    # 确保进度条显示100%
                    self.progress_var.set(100)
                    messagebox.showinfo("成功", msg_data)
                elif msg_type == "error":
                    self.log_message(msg_data)
                    # 错误时重置进度条
                    self.progress_var.set(0)
                    messagebox.showerror("错误", msg_data)
                elif msg_type == "finished":
                    self.start_btn.config(state="normal")
                    self.cancel_btn.config(state="disabled")
                    self.is_exporting = False
                    # 如果没有错误，确保进度条显示完成
                    if self.progress_var.get() > 0:
                        self.progress_var.set(100)
                
        except queue.Empty:
            pass
        
        # 每100ms检查一次队列
        self.root.after(100, self.check_queue)
    
    def log_message(self, message):
        """记录消息到日志区域"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.result_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.result_text.see(tk.END)

def main():
    root = tk.Tk()
    app = MailExporterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()