import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import threading
import queue
import os
from mail_exporter import fetch_emails, get_mail_folders, get_supported_providers, fetch_emails_incremental

# 尝试导入日期选择器，如果没有安装则使用普通输入框
try:
    from tkcalendar import DateEntry
    HAS_DATE_PICKER = True
except ImportError:
    HAS_DATE_PICKER = False

# 程序版本信息
VERSION = "1.5.0"

class MailExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("邮箱邮件导出工具")
        # 设置窗口为最大化显示
        self.root.state('zoomed')
        self.root.resizable(True, True)
        # 设置最小窗口尺寸
        self.root.minsize(650, 600)
        
        # 创建消息队列用于线程间通信
        self.message_queue = queue.Queue()
        self.export_thread = None
        self.is_exporting = False
        self.stop_requested = False
        self.folder_mapping = {}  # 文件夹显示名到实际名的映射
        
        self.setup_ui()
        self.check_queue()
        
        # 添加欢迎消息到日志框
        self.log_message("邮件导出工具已启动，请配置参数后开始导出。")
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        # 确保结果区域（第5行）有足够的权重，防止被压缩
        main_frame.rowconfigure(5, weight=1)
        
        # 标题和版本信息
        title_label = ttk.Label(main_frame, text="邮箱邮件导出工具", font=('Arial', 16, 'bold'))
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
        
        # 初始化时检查OAuth状态（延迟执行，确保UI完全创建）
        self.root.after(100, self._check_oauth_status)
    
    def create_input_fields(self, parent):
        """创建输入字段"""
        # 输入字段框架
        input_frame = ttk.LabelFrame(parent, text="邮箱配置", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # 邮箱服务提供商
        ttk.Label(input_frame, text="邮箱服务商:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.provider_var = tk.StringVar()
        provider_frame = ttk.Frame(input_frame)
        provider_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        provider_frame.columnconfigure(0, weight=1)
        
        # 获取支持的邮箱服务提供商
        try:
            providers = get_supported_providers()
            provider_values = [f"{name} ({display})" for name, display in providers.items()]
        except:
            provider_values = ["163 (网易邮箱)", "gmail (Gmail)", "qq (QQ邮箱)", "outlook (Outlook)", "yahoo (Yahoo)", "custom (其他邮箱(自定义))"]
        
        self.provider_combobox = ttk.Combobox(provider_frame, textvariable=self.provider_var,
                                            values=provider_values, state="readonly", width=25)
        self.provider_combobox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.provider_combobox.set(provider_values[0])  # 默认选择第一个
        row += 1

        # 自定义邮箱配置框架（仅在选择其他邮箱时显示）
        self.custom_provider_frame = ttk.LabelFrame(input_frame, text="自定义邮箱服务器配置", padding="5")
        self.custom_provider_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        self.custom_provider_frame.columnconfigure(1, weight=1)

        # IMAP服务器地址
        ttk.Label(self.custom_provider_frame, text="IMAP服务器:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.custom_imap_server_var = tk.StringVar()
        custom_server_entry = ttk.Entry(self.custom_provider_frame, textvariable=self.custom_imap_server_var, width=40)
        custom_server_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)

        # IMAP端口
        ttk.Label(self.custom_provider_frame, text="IMAP端口:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.custom_imap_port_var = tk.StringVar(value="993")
        custom_port_entry = ttk.Entry(self.custom_provider_frame, textvariable=self.custom_imap_port_var, width=10)
        custom_port_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=2)

        # 是否使用SSL
        self.custom_use_ssl_var = tk.BooleanVar(value=True)
        custom_ssl_checkbox = ttk.Checkbutton(self.custom_provider_frame, text="使用SSL/TLS加密连接",
                                             variable=self.custom_use_ssl_var)
        custom_ssl_checkbox.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)

        # 说明文本
        custom_note = ttk.Label(self.custom_provider_frame,
                               text="例如: imap.qq.com (端口993), imap.sina.com (端口993)等",
                               font=('Arial', 8), foreground='gray')
        custom_note.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))

        # 初始隐藏自定义配置框
        self.custom_provider_frame.grid_remove()

        row += 1

        # 代理设置
        ttk.Label(input_frame, text="代理设置:").grid(row=row, column=0, sticky=tk.W, pady=2)
        proxy_frame = ttk.Frame(input_frame)
        proxy_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        proxy_frame.columnconfigure(1, weight=1)
        
        # 代理启用选项
        self.use_proxy_var = tk.BooleanVar(value=False)
        proxy_checkbox = ttk.Checkbutton(proxy_frame, text="启用代理连接", 
                                       variable=self.use_proxy_var,
                                       command=self.toggle_proxy_config)
        proxy_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        # 代理类型选择
        proxy_type_frame = ttk.Frame(proxy_frame)
        proxy_type_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(20, 0))
        
        ttk.Label(proxy_type_frame, text="类型:").grid(row=0, column=0, sticky=tk.W)
        self.proxy_type_var = tk.StringVar(value="SOCKS5")
        proxy_type_combo = ttk.Combobox(proxy_type_frame, textvariable=self.proxy_type_var,
                                      values=["SOCKS5", "SOCKS4", "HTTP"], state="readonly", width=8)
        proxy_type_combo.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        row += 1
        
        # 代理服务器配置框架
        self.proxy_config_frame = ttk.Frame(input_frame)
        self.proxy_config_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        self.proxy_config_frame.columnconfigure(1, weight=1)
        self.proxy_config_frame.columnconfigure(3, weight=1)
        
        # 代理服务器地址
        ttk.Label(self.proxy_config_frame, text="代理服务器:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.proxy_host_var = tk.StringVar(value="127.0.0.1")
        proxy_host_entry = ttk.Entry(self.proxy_config_frame, textvariable=self.proxy_host_var, width=20)
        proxy_host_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=2)
        
        # 代理端口
        ttk.Label(self.proxy_config_frame, text="端口:").grid(row=0, column=2, sticky=tk.W, pady=2, padx=(10, 0))
        self.proxy_port_var = tk.StringVar(value="7890")
        proxy_port_entry = ttk.Entry(self.proxy_config_frame, textvariable=self.proxy_port_var, width=8)
        proxy_port_entry.grid(row=0, column=3, sticky=tk.W, padx=(5, 0), pady=2)
        
        # 代理认证
        auth_frame = ttk.Frame(self.proxy_config_frame)
        auth_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(5, 0))
        auth_frame.columnconfigure(1, weight=1)
        auth_frame.columnconfigure(3, weight=1)
        
        self.proxy_auth_var = tk.BooleanVar(value=False)
        auth_checkbox = ttk.Checkbutton(auth_frame, text="需要认证", 
                                      variable=self.proxy_auth_var,
                                      command=self.toggle_proxy_auth)
        auth_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        # 代理用户名
        self.proxy_username_label = ttk.Label(auth_frame, text="用户名:")
        self.proxy_username_label.grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        self.proxy_username_var = tk.StringVar()
        self.proxy_username_entry = ttk.Entry(auth_frame, textvariable=self.proxy_username_var, width=15)
        self.proxy_username_entry.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(5, 10))
        
        # 代理密码
        self.proxy_password_label = ttk.Label(auth_frame, text="密码:")
        self.proxy_password_label.grid(row=0, column=3, sticky=tk.W)
        self.proxy_password_var = tk.StringVar()
        self.proxy_password_entry = ttk.Entry(auth_frame, textvariable=self.proxy_password_var, show="*", width=15)
        self.proxy_password_entry.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(5, 0))
        
        # 初始状态设置
        self.toggle_proxy_config()
        self.toggle_proxy_auth()
        
        row += 1
        
        # 用户名
        ttk.Label(input_frame, text="邮箱用户名:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(input_frame, textvariable=self.username_var, width=30)
        username_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        row += 1
        
        # 密码
        ttk.Label(input_frame, text="密码/授权码:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(input_frame, textvariable=self.password_var, show="*", width=30)
        self.password_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        row += 1
        
        # OAuth配置（仅Gmail显示）
        self.oauth_frame = ttk.Frame(input_frame)
        self.oauth_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        self.oauth_frame.columnconfigure(1, weight=1)
        
        # OAuth启用选项
        self.use_oauth_var = tk.BooleanVar(value=False)
        oauth_checkbox = ttk.Checkbutton(self.oauth_frame, text="使用OAuth2认证（推荐）", 
                                       variable=self.use_oauth_var,
                                       command=self.toggle_oauth_config)
        oauth_checkbox.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # OAuth配置字段
        self.oauth_config_frame = ttk.LabelFrame(self.oauth_frame, text="OAuth2配置", padding="5")
        self.oauth_config_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        self.oauth_config_frame.columnconfigure(1, weight=1)
        
        # Client ID
        self.client_id_label = ttk.Label(self.oauth_config_frame, text="Client ID:")
        self.client_id_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        self.client_id_var = tk.StringVar()
        self.client_id_entry = ttk.Entry(self.oauth_config_frame, textvariable=self.client_id_var, width=50)
        self.client_id_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # Client Secret
        self.client_secret_label = ttk.Label(self.oauth_config_frame, text="Client Secret:")
        self.client_secret_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        self.client_secret_var = tk.StringVar()
        self.client_secret_entry = ttk.Entry(self.oauth_config_frame, textvariable=self.client_secret_var, show="*", width=50)
        self.client_secret_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        # OAuth授权按钮和状态
        oauth_auth_frame = ttk.Frame(self.oauth_config_frame)
        oauth_auth_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5))
        oauth_auth_frame.columnconfigure(1, weight=1)
        
        self.oauth_auth_button = ttk.Button(oauth_auth_frame, text="开始OAuth授权", command=self.start_oauth_auth)
        self.oauth_auth_button.grid(row=0, column=0, sticky=tk.W)
        
        self.oauth_status_var = tk.StringVar(value="未授权")
        self.oauth_status_label = ttk.Label(oauth_auth_frame, textvariable=self.oauth_status_var, foreground="red")
        self.oauth_status_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # OAuth状态说明
        self.oauth_status_note_var = tk.StringVar(value="需要填入Client ID和Secret进行授权")
        self.oauth_status_note = ttk.Label(self.oauth_config_frame, textvariable=self.oauth_status_note_var, 
                                         font=('Arial', 8), foreground='gray')
        self.oauth_status_note.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # 说明文本
        oauth_note = ttk.Label(self.oauth_config_frame, 
                             text="请在Google Cloud Console创建OAuth2凭据并填入上述信息，然后点击授权按钮", 
                             font=('Arial', 8), foreground='blue')
        oauth_note.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # 绑定提供商变化事件
        self.provider_combobox.bind('<<ComboboxSelected>>', self.on_provider_changed)
        
        # 初始状态设置
        self.toggle_oauth_config()
        
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
        
        # 邮件数量限制
        ttk.Label(input_frame, text="邮件数量:").grid(row=row, column=0, sticky=tk.W, pady=2)
        count_frame = ttk.Frame(input_frame)
        count_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=2)
        
        self.email_count_var = tk.StringVar(value="100")
        count_entry = ttk.Entry(count_frame, textvariable=self.email_count_var, width=10)
        count_entry.grid(row=0, column=0)
        
        ttk.Label(count_frame, text="封 (0表示不限制)").grid(row=0, column=1, padx=(5, 0))
        
        # 添加说明文本
        count_note = ttk.Label(input_frame, text="(获取最近的N封邮件，按时间倒序)", 
                              font=('Arial', 8), foreground='gray')
        count_note.grid(row=row+1, column=1, sticky=tk.W, padx=(10, 0), pady=(0, 5))
        row += 2
        
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
        
        self.cancel_btn = ttk.Button(button_frame, text="停止", command=self.cancel_export, state="disabled")
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
    
    def refresh_folders(self):
        """刷新邮箱文件夹列表"""
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        provider_text = self.provider_var.get().strip()
        
        if not username:
            messagebox.showwarning("提示", "请先输入邮箱用户名")
            return
        
        if not provider_text:
            messagebox.showwarning("提示", "请选择邮箱服务商")
            return
        
        # 检查认证方式
        provider = provider_text.split(' (')[0] if ' (' in provider_text else provider_text
        use_oauth = self.use_oauth_var.get() and 'gmail' in provider.lower()
        
        if not use_oauth and not password:
            messagebox.showwarning("提示", "请先输入密码或授权码")
            return
        
        # 从显示文本中提取provider名称
        provider = provider_text.split(' (')[0] if ' (' in provider_text else provider_text
        
        # 显示加载状态
        original_text = self.folder_combobox['state']
        self.folder_combobox.config(state="disabled")
        self.status_var.set("正在获取文件夹列表...")
        
        def get_folders_worker():
            try:
                # 获取代理设置
                proxy_config = None
                if self.use_proxy_var.get():
                    proxy_config = {
                        'enabled': True,
                        'type': self.proxy_type_var.get(),
                        'host': self.proxy_host_var.get(),
                        'port': int(self.proxy_port_var.get()) if self.proxy_port_var.get() else None,
                        'username': self.proxy_username_var.get() if self.proxy_username_var.get() else None,
                        'password': self.proxy_password_var.get() if self.proxy_password_var.get() else None
                    }
                
                # 获取OAuth配置
                oauth_config = None
                if self.use_oauth_var.get() and 'gmail' in provider.lower():
                    oauth_config = {
                        'client_id': self.client_id_var.get().strip(),
                        'client_secret': self.client_secret_var.get().strip(),
                        'credentials_file': None,
                        'token_file': 'gmail_token.json'
                    }
                
                folders = get_mail_folders(username, password, provider, proxy_config, oauth_config)
                
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
    
    def toggle_proxy_config(self):
        """切换代理配置界面显示状态"""
        if self.use_proxy_var.get():
            # 显示代理配置框架
            self.proxy_config_frame.grid()
            # 显示代理配置内容
            for widget in self.proxy_config_frame.winfo_children():
                widget.grid()
            # 递归显示子框架中的控件
            for child in self.proxy_config_frame.winfo_children():
                if isinstance(child, ttk.Frame):
                    for subwidget in child.winfo_children():
                        subwidget.grid()
        else:
            # 完全隐藏代理配置框架
            self.proxy_config_frame.grid_remove()
        
        # 强制更新布局并调整窗口大小
        self.root.update_idletasks()
        
        # 获取当前窗口的最小尺寸
        self.root.update()
        min_width = self.root.winfo_reqwidth()
        min_height = self.root.winfo_reqheight()
        
        # 如果隐藏代理配置，调整窗口大小
        if not self.use_proxy_var.get():
            current_width = self.root.winfo_width()
            # 保持当前宽度，但调整高度以适应内容
            self.root.geometry(f"{max(current_width, min_width)}x{min_height}")
        
        # 设置最小窗口大小
        self.root.minsize(650, 600)
    
    def toggle_proxy_auth(self):
        """切换代理认证配置显示"""
        if self.proxy_auth_var.get():
            # 显示认证字段
            self.proxy_username_label.grid()
            self.proxy_username_entry.grid()
            self.proxy_password_label.grid()
            self.proxy_password_entry.grid()
        else:
            # 隐藏认证字段
            self.proxy_username_label.grid_remove()
            self.proxy_username_entry.grid_remove()
            self.proxy_password_label.grid_remove()
            self.proxy_password_entry.grid_remove()
    
    def toggle_oauth_config(self):
        """切换OAuth配置显示"""
        if self.use_oauth_var.get():
            # 显示OAuth配置
            self.oauth_config_frame.grid()
            # 禁用密码输入
            self.password_entry.config(state='disabled')
        else:
            # 隐藏OAuth配置
            self.oauth_config_frame.grid_remove()
            # 启用密码输入
            self.password_entry.config(state='normal')
    
    def check_oauth_status(self):
        """检查OAuth授权状态"""
        self._check_oauth_status()
    
    def refresh_oauth_status(self):
        """强制刷新OAuth授权状态"""
        self.log_message("正在刷新OAuth状态...")
        self._check_oauth_status()
        self.log_message("OAuth状态刷新完成")
    
    def _hide_oauth_credentials(self):
        """隐藏OAuth凭据字段"""
        try:
            self.client_id_label.grid_remove()
            self.client_id_entry.grid_remove()
            self.client_secret_label.grid_remove()
            self.client_secret_entry.grid_remove()
        except AttributeError:
            # 如果控件还未创建，忽略错误
            pass
    
    def _show_oauth_credentials(self):
        """显示OAuth凭据字段"""
        try:
            self.client_id_label.grid()
            self.client_id_entry.grid()
            self.client_secret_label.grid()
            self.client_secret_entry.grid()
        except AttributeError:
            # 如果控件还未创建，忽略错误
            pass
    
    def _check_oauth_status(self):
        """检查OAuth授权状态"""
        try:
            if os.path.exists('gmail_token.json'):
                from oauth_gmail import GmailOAuth
                
                # 尝试使用当前的Client ID和Secret创建OAuth实例
                client_id = self.client_id_var.get().strip() if hasattr(self, 'client_id_var') else ""
                client_secret = self.client_secret_var.get().strip() if hasattr(self, 'client_secret_var') else ""
                
                # 如果有凭据，使用它们创建OAuth实例
                if client_id and client_secret:
                    gmail_oauth = GmailOAuth(
                        client_id=client_id,
                        client_secret=client_secret,
                        token_file='gmail_token.json'
                    )
                else:
                    # 如果没有凭据，尝试从令牌文件中读取
                    gmail_oauth = GmailOAuth(token_file='gmail_token.json')
                
                # 检查令牌是否有效
                if gmail_oauth._load_existing_token() and gmail_oauth.credentials:
                    # 检查令牌是否过期
                    if gmail_oauth.credentials.valid:
                        # OAuth已授权且令牌有效
                        self.oauth_status_var.set("已授权")
                        self.oauth_status_label.config(foreground="green")
                        self.oauth_auth_button.config(text="重新授权")
                        self.oauth_status_note_var.set("OAuth2已授权，可直接导出邮件。如需重新授权请点击重新授权按钮")
                        self._hide_oauth_credentials()
                        self.log_message("OAuth状态检查: 已授权")
                        return True
                    elif gmail_oauth.credentials.expired and gmail_oauth.credentials.refresh_token:
                        # 令牌过期但有刷新令牌，尝试刷新
                        try:
                            if gmail_oauth._refresh_token():
                                self.oauth_status_var.set("已授权")
                                self.oauth_status_label.config(foreground="green")
                                self.oauth_auth_button.config(text="重新授权")
                                self.oauth_status_note_var.set("OAuth2已授权，可直接导出邮件。如需重新授权请点击重新授权按钮")
                                self._hide_oauth_credentials()
                                self.log_message("OAuth状态检查: 令牌已刷新，已授权")
                                return True
                            else:
                                self.log_message("OAuth状态检查: 令牌刷新失败")
                        except Exception as refresh_error:
                            self.log_message(f"OAuth状态检查: 令牌刷新出错 - {refresh_error}")
                    else:
                        # 令牌过期且无刷新令牌
                        self.log_message("OAuth状态检查: 令牌已过期且无刷新令牌")
                else:
                    # 令牌文件存在但无效
                    self.log_message("OAuth状态检查: 令牌无效")
        except Exception as e:
            self.log_message(f"OAuth状态检查出错: {e}")
        
        # OAuth未授权或检查失败
        self.oauth_status_var.set("未授权")
        self.oauth_status_label.config(foreground="red")
        self.oauth_auth_button.config(text="开始OAuth授权")
        self.oauth_status_note_var.set("需要填入Client ID和Secret进行授权")
        self._show_oauth_credentials()
        self.log_message("OAuth状态检查: 未授权")
        return False
    
    def start_oauth_auth(self):
        """开始OAuth授权流程"""
        # 如果是重新授权，先显示凭据字段
        if self.oauth_auth_button.cget("text") == "重新授权":
            self._show_oauth_credentials()
            self.oauth_status_var.set("请填入凭据后重新授权")
            self.oauth_status_label.config(foreground="orange")
            self.oauth_status_note_var.set("请填入Client ID和Secret，然后点击开始OAuth授权")
            self.oauth_auth_button.config(text="开始OAuth授权")
            return
        
        client_id = self.client_id_var.get().strip()
        client_secret = self.client_secret_var.get().strip()
        
        if not client_id or not client_secret:
            messagebox.showerror("错误", "请先填入Client ID和Client Secret")
            return
        
        self.oauth_auth_button.config(state="disabled")
        self.oauth_status_var.set("授权中...")
        self.oauth_status_label.config(foreground="orange")
        
        def oauth_worker():
             try:
                 from oauth_gmail import GmailOAuth
                 
                 def status_callback(message):
                     """状态回调函数，在主线程中更新状态"""
                     # 避免在认证完成后覆盖最终状态
                     if not message.startswith("OAuth授权完成"):
                         self.root.after(0, lambda msg=message: self.update_oauth_status(msg))
                 
                 gmail_oauth = GmailOAuth(
                     client_id=client_id,
                     client_secret=client_secret,
                     credentials_file=None,
                     token_file='gmail_token.json',
                     status_callback=status_callback
                 )
                 
                 # 执行OAuth认证，强制重新认证以确保获取最新授权
                 success = gmail_oauth.authenticate(force_reauth=True)
                 
                 if success:
                     # 确保成功状态在最后设置
                     self.root.after(0, lambda: self.oauth_auth_success())
                 else:
                     self.root.after(0, lambda: self.oauth_auth_failed("授权失败"))
                     
             except Exception as e:
                 error_msg = str(e)
                 self.root.after(0, lambda msg=error_msg: self.oauth_auth_failed(msg))
        
        threading.Thread(target=oauth_worker, daemon=True).start()
    
    def update_oauth_status(self, message):
        """更新OAuth状态显示"""
        self.oauth_status_var.set(message)
        self.log_message(f"OAuth状态: {message}")
    
    def oauth_auth_success(self):
        """OAuth授权成功"""
        # 明确设置成功状态
        self.oauth_status_var.set("已授权")
        self.oauth_status_label.config(foreground="green")
        self.oauth_auth_button.config(state="normal", text="重新授权")
        self.oauth_status_note_var.set("OAuth2已授权，可直接导出邮件。如需重新授权请点击重新授权按钮")
        
        # 隐藏OAuth凭据字段
        self._hide_oauth_credentials()
        
        self.log_message("OAuth授权成功")
        
        # 延迟检查OAuth状态以确保令牌文件已保存
        self.root.after(1000, self._check_oauth_status)
    
    def oauth_auth_failed(self, error_msg):
        """OAuth授权失败"""
        self.oauth_status_var.set("授权失败")
        self.oauth_status_label.config(foreground="red")
        self.oauth_auth_button.config(state="normal")
        
        # 根据错误类型提供不同的提示信息
        if "授权码无效" in error_msg or "invalid_grant" in error_msg:
            note_msg = "授权码已过期或无效，请重新进行OAuth授权"
            dialog_msg = "授权码已过期或无效。\n\n这通常发生在：\n1. 授权过程中等待时间过长\n2. 重复使用了已失效的授权码\n\n请点击'开始OAuth授权'重新进行授权。"
        elif "权限范围已更改" in error_msg or "范围冲突" in error_msg or "Scope has changed" in error_msg:
            note_msg = "OAuth权限范围已更改，请重新授权"
            dialog_msg = "OAuth权限范围已更改。\n\n应用的权限要求发生了变化，旧的授权令牌已失效。\n\n解决方案：\n1. 点击'开始OAuth授权'重新进行授权\n2. 在浏览器中重新确认权限\n3. 完成授权后即可正常使用"
        elif "Client ID" in error_msg or "client_secret" in error_msg:
            note_msg = "Client ID或Secret无效，请检查配置"
            dialog_msg = "Client ID或Client Secret无效。\n\n请检查：\n1. Client ID和Secret是否正确\n2. 是否已在Google Cloud Console中启用Gmail API\n3. 重定向URI是否正确配置"
        elif "超时" in error_msg or "timeout" in error_msg.lower():
            note_msg = "授权超时，请重新尝试"
            dialog_msg = "OAuth授权超时。\n\n请确保：\n1. 网络连接正常\n2. 在5分钟内完成浏览器授权\n3. 没有被防火墙阻止"
        else:
            note_msg = "授权失败，请检查网络连接和配置"
            dialog_msg = f"OAuth授权失败：\n\n{error_msg}\n\n请检查网络连接和OAuth配置是否正确。"
        
        self.oauth_status_note_var.set(note_msg)
        # OAuth授权失败后显示凭据字段
        self._show_oauth_credentials()
        messagebox.showerror("OAuth授权失败", dialog_msg)
        self.log_message(f"OAuth授权失败: {error_msg}")
        
        # 刷新OAuth状态以确保界面状态正确
        self.root.after(500, self._check_oauth_status)
    
    def on_provider_changed(self, event=None):
        """当邮箱提供商改变时的处理"""
        provider_text = self.provider_var.get()
        provider = provider_text.split(' (')[0].strip().lower() if ' (' in provider_text else provider_text.lower()

        # 处理Gmail OAuth选项显示
        if 'gmail' in provider:
            # 显示OAuth选项
            self.oauth_frame.grid()
            self.custom_provider_frame.grid_remove()
        elif provider == 'custom':
            # 隐藏OAuth选项，显示自定义配置
            self.oauth_frame.grid_remove()
            self.use_oauth_var.set(False)
            self.toggle_oauth_config()
            self.custom_provider_frame.grid()
        else:
            # 隐藏OAuth选项和自定义配置
            self.oauth_frame.grid_remove()
            self.use_oauth_var.set(False)
            self.toggle_oauth_config()
            self.custom_provider_frame.grid_remove()
    
    def start_export(self):
        """开始导出"""
        # 验证输入
        if not self.validate_inputs():
            return
        
        # 禁用开始按钮，启用停止按钮
        self.start_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.is_exporting = True
        
        # 清空结果文本
        self.result_text.delete(1.0, tk.END)
        
        # 重置进度和停止标志
        self.progress_var.set(0)
        self.stop_requested = False
        self.status_var.set("正在连接邮箱服务器...")
        
        # 启动导出线程
        self.export_thread = threading.Thread(target=self.export_worker, daemon=True)
        self.export_thread.start()
    
    def cancel_export(self):
        """停止导出"""
        self.stop_requested = True
        self.log_message("正在停止导出，请稍候...")
        self.status_var.set("正在停止...")
    
    def validate_inputs(self):
        """验证输入"""
        if not self.username_var.get().strip():
            messagebox.showerror("错误", "请输入邮箱用户名")
            return False

        # 获取提供商信息
        provider_text = self.provider_var.get().strip()
        provider = provider_text.split(' (')[0].strip().lower() if ' (' in provider_text else provider_text.lower()

        # 验证自定义邮箱配置
        if provider == 'custom':
            imap_server = self.custom_imap_server_var.get().strip()
            imap_port_str = self.custom_imap_port_var.get().strip()

            if not imap_server:
                messagebox.showerror("错误", "请输入IMAP服务器地址")
                return False

            if not imap_port_str:
                messagebox.showerror("错误", "请输入IMAP端口")
                return False

            try:
                imap_port = int(imap_port_str)
                if imap_port <= 0 or imap_port > 65535:
                    messagebox.showerror("错误", "IMAP端口必须在1-65535之间")
                    return False
            except ValueError:
                messagebox.showerror("错误", "IMAP端口必须是数字")
                return False

        # 检查认证方式
        use_oauth = self.use_oauth_var.get() and 'gmail' in provider

        if use_oauth:
            # OAuth认证验证 - 先检查是否已有有效令牌
            oauth_authorized = False
            try:
                if os.path.exists('gmail_token.json'):
                    from oauth_gmail import GmailOAuth
                    gmail_oauth = GmailOAuth(token_file='gmail_token.json')
                    if gmail_oauth._load_existing_token() and gmail_oauth.credentials and gmail_oauth.credentials.valid:
                        oauth_authorized = True
            except:
                pass

            # 如果OAuth未授权，则需要检查Client ID和Secret
            if not oauth_authorized:
                if not self.client_id_var.get().strip():
                    messagebox.showerror("错误", "请输入OAuth Client ID或先完成OAuth授权")
                    return False
                if not self.client_secret_var.get().strip():
                    messagebox.showerror("错误", "请输入OAuth Client Secret或先完成OAuth授权")
                    return False
        else:
            # 传统密码认证验证
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
        
        # 验证邮件数量
        email_count_str = self.email_count_var.get().strip()
        if email_count_str:
            try:
                email_count = int(email_count_str)
                if email_count < 0:
                    messagebox.showerror("错误", "邮件数量不能为负数")
                    return False
            except ValueError:
                messagebox.showerror("错误", "邮件数量必须是数字")
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
            provider_text = self.provider_var.get().strip()
            
            # 从显示文本中提取provider名称
            provider = provider_text.split(' (')[0] if ' (' in provider_text else provider_text
            
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
            
            # 获取邮件数量限制
            email_count_str = self.email_count_var.get().strip()
            email_count_limit = 0  # 默认不限制
            if email_count_str:
                try:
                    email_count_limit = int(email_count_str)
                except ValueError:
                    email_count_limit = 0
            
            # 发送开始消息
            self.message_queue.put(("log", f"开始导出邮件..."))
            self.message_queue.put(("log", f"邮箱服务商: {provider}"))
            self.message_queue.put(("log", f"用户: {username}"))
            self.message_queue.put(("log", f"时间范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}"))
            self.message_queue.put(("log", f"输出文件: {output_file}"))
            self.message_queue.put(("log", f"邮箱文件夹: {folder}"))
            if email_count_limit > 0:
                self.message_queue.put(("log", f"邮件数量限制: 最近 {email_count_limit} 封"))
            else:
                self.message_queue.put(("log", "邮件数量限制: 无限制"))
            
            if download_attachments:
                if attachment_folder:
                    self.message_queue.put(("log", f"附件保存目录: {attachment_folder}"))
                else:
                    self.message_queue.put(("log", "附件保存目录: CSV文件同目录下的attachments文件夹"))
            else:
                self.message_queue.put(("log", "不下载附件"))
            
            # 获取代理设置
            use_proxy = self.use_proxy_var.get()
            proxy_config = None
            
            if use_proxy:
                proxy_config = {
                    'type': self.proxy_type_var.get().lower(),
                    'host': self.proxy_host_var.get().strip(),
                    'port': int(self.proxy_port_var.get().strip()) if self.proxy_port_var.get().strip().isdigit() else 1080,
                    'username': self.proxy_username_var.get().strip() if self.proxy_auth_var.get() else None,
                    'password': self.proxy_password_var.get().strip() if self.proxy_auth_var.get() else None
                }
                self.message_queue.put(("log", f"代理设置: {proxy_config['type'].upper()} {proxy_config['host']}:{proxy_config['port']}"))
                if proxy_config['username']:
                    self.message_queue.put(("log", f"代理认证: 用户名 {proxy_config['username']}"))
            
            # 获取OAuth配置
            oauth_config = None
            use_oauth = self.use_oauth_var.get() and 'gmail' in provider

            if use_oauth:
                oauth_config = {
                    'client_id': self.client_id_var.get().strip(),
                    'client_secret': self.client_secret_var.get().strip(),
                    'credentials_file': None,  # 使用内置凭据
                    'token_file': 'gmail_token.json'
                }
                self.message_queue.put(("log", "使用OAuth2认证"))
                # OAuth模式下密码参数可以为空
                password = None

            # 获取自定义邮箱参数
            custom_imap_server = None
            custom_imap_port = None
            custom_use_ssl = True

            if provider == 'custom':
                custom_imap_server = self.custom_imap_server_var.get().strip()
                try:
                    custom_imap_port = int(self.custom_imap_port_var.get().strip())
                except ValueError:
                    custom_imap_port = 993
                custom_use_ssl = self.custom_use_ssl_var.get()
                self.message_queue.put(("log", f"自定义IMAP服务器: {custom_imap_server}:{custom_imap_port}"))
                if custom_use_ssl:
                    self.message_queue.put(("log", "使用SSL/TLS加密连接"))

            # 调用增量导出函数，传入进度回调、附件参数、停止标志、代理设置、OAuth配置和邮件数量限制
            email_count = fetch_emails_incremental(username, password, start_date, end_date, output_file, folder,
                                                  self.progress_callback, download_attachments, attachment_folder,
                                                  provider, lambda: self.stop_requested, proxy_config, oauth_config, email_count_limit,
                                                  custom_imap_server, custom_imap_port, custom_use_ssl)
            
            if self.is_exporting:
                if self.stop_requested:
                    self.message_queue.put(("success", f"导出已停止! 共导出 {email_count} 封邮件到 {output_file}"))
                    self.message_queue.put(("status", "已停止"))
                else:
                    self.message_queue.put(("success", f"导出完成! 共导出 {email_count} 封邮件到 {output_file}"))
                    self.message_queue.put(("progress", 100))
                    self.message_queue.put(("status", "导出完成"))
            
        except Exception as e:
            if self.is_exporting:
                if self.stop_requested:
                    self.message_queue.put(("error", f"导出已停止: {str(e)}"))
                    self.message_queue.put(("status", "已停止"))
                else:
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
                    self.stop_requested = False
                    # 如果没有错误且未被停止，确保进度条显示完成
                    if self.progress_var.get() > 0 and not self.stop_requested:
                        self.progress_var.set(100)
                
        except queue.Empty:
            pass
        
        # 每100ms检查一次队列
        self.root.after(100, self.check_queue)
    
    def log_message(self, message):
        """记录消息到日志区域"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        # 检查result_text是否存在，避免在UI初始化期间出错
        if hasattr(self, 'result_text') and self.result_text:
            self.result_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.result_text.see(tk.END)
        else:
            # 如果UI还未完全初始化，只打印到控制台
            print(f"[{timestamp}] {message}")

def main():
    root = tk.Tk()
    app = MailExporterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()