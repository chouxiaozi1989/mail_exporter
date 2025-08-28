#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件客户端模块
负责连接和操作不同的邮箱服务提供商
"""

import imaplib
import email
import time
import locale
from datetime import datetime, timedelta
from typing import List, Tuple, Optional, Dict, Any
from email_providers import EmailProviders, EmailProviderConfig
from email_parser import EmailParser
from email_exporter import EmailExporter
from incremental_exporter import IncrementalEmailExporter


class MailClient:
    """邮件客户端类"""
    
    def __init__(self, provider_name: str = None, email_address: str = None):
        """
        初始化邮件客户端
        
        Args:
            provider_name: 邮箱服务提供商名称
            email_address: 邮箱地址（用于自动检测提供商）
        """
        self.providers = EmailProviders()
        self.parser = EmailParser()
        self.exporter = EmailExporter()
        self.mail = None
        self.config = None
        
        # 确定邮箱配置
        if provider_name:
            self.config = self.providers.get_provider_by_name(provider_name)
        elif email_address:
            self.config = self.providers.get_provider_by_email(email_address)
        
        if not self.config:
            raise ValueError("无法确定邮箱服务提供商配置")
    
    def connect(self, username: str, password: str) -> bool:
        """
        连接到邮箱服务器
        
        Args:
            username: 用户名
            password: 密码或授权码
            
        Returns:
            连接是否成功
        """
        try:
            # 验证邮箱地址格式
            if not self.providers.validate_email_format(username):
                raise ValueError(f"邮箱地址格式不正确")
            
            # 获取连接参数
            conn_params = self.providers.get_connection_params(username, self.config.name)
            
            # 建立IMAP连接
            if self.config.imap_port == 993:
                self.mail = imaplib.IMAP4_SSL(host=self.config.imap_server, port=self.config.imap_port)
            else:
                self.mail = imaplib.IMAP4(host=self.config.imap_server, port=self.config.imap_port)
                if self.config.use_ssl:
                    self.mail.starttls()
            
            # 登录
            self.mail.login(username, password)
            
            # 发送客户端标识（可选，某些服务器需要）
            try:
                if hasattr(imaplib, 'Commands') and "ID" not in imaplib.Commands:
                    imaplib.Commands["ID"] = ('AUTH',)
                args = ("name", username, "contact", username, "version", "1.0.0", "vendor", "mail_exporter")
                self.mail._simple_command("ID", str(args).replace(",", "").replace("'", '"'))
            except:
                pass  # 忽略ID命令失败
            
            return True
            
        except Exception as e:
            raise Exception(f"连接到 {self.config.name} 失败: {str(e)}")
    
    def disconnect(self):
        """断开邮箱连接"""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
            except:
                pass
            finally:
                self.mail = None
    
    def get_folders(self) -> List[Tuple[str, str]]:
        """
        获取邮箱文件夹列表
        
        Returns:
            文件夹列表，格式为 [(folder_name, display_name), ...]
        """
        if not self.mail:
            raise Exception("未连接到邮箱服务器")
        
        folders = []
        try:
            # 获取所有文件夹
            status, folder_list = self.mail.list()
            
            if status == 'OK':
                for folder_info in folder_list:
                    # 解析文件夹信息
                    folder_str = None
                    for encoding in ['utf-8', 'gb2312', 'gbk', 'latin1']:
                        try:
                            folder_str = folder_info.decode(encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    
                    if folder_str is None:
                        folder_str = folder_info.decode('utf-8', errors='ignore')
                    
                    # 提取文件夹名称
                    import re
                    match = re.search(r'"([^"]+)"\s*$', folder_str)
                    if match:
                        folder_name = match.group(1)
                        
                        # 解码IMAP UTF-7编码
                        if '&' in folder_name:
                            try:
                                folder_name = self.parser.decode_imap_utf7(folder_name)
                            except:
                                pass
                        
                        # 创建显示名称
                        display_name = self._get_folder_display_name(folder_name)
                        folders.append((folder_name, display_name))
            
            # 确保INBOX在第一位
            folders.sort(key=lambda x: (x[0] != 'INBOX', x[1]))
            
        except Exception as e:
            raise Exception(f"获取文件夹列表失败: {str(e)}")
        
        return folders
    
    def _get_folder_display_name(self, folder_name: str) -> str:
        """
        获取文件夹的显示名称
        
        Args:
            folder_name: 原始文件夹名称
            
        Returns:
            本地化的显示名称
        """
        # 通用文件夹映射
        folder_mapping = {
            'INBOX': '收件箱',
            'Sent Messages': '已发送',
            'Sent': '已发送',
            'Drafts': '草稿箱',
            'Deleted Messages': '已删除',
            'Trash': '垃圾箱',
            'Junk': '垃圾邮件',
            'Spam': '垃圾邮件',
            'Outbox': '发件箱'
        }
        
        # Gmail特定文件夹
        if self.config.name.lower() == 'gmail':
            gmail_mapping = {
                '[Gmail]/Sent Mail': '已发送',
                '[Gmail]/Drafts': '草稿箱',
                '[Gmail]/Trash': '垃圾箱',
                '[Gmail]/Spam': '垃圾邮件',
                '[Gmail]/All Mail': '所有邮件',
                '[Gmail]/Starred': '已加星标',
                '[Gmail]/Important': '重要邮件'
            }
            folder_mapping.update(gmail_mapping)
        
        # QQ邮箱特定文件夹
        elif self.config.name.lower() == 'qq':
            qq_mapping = {
                'Sent Messages': '已发送',
                'Deleted Messages': '已删除',
                'Junk': '垃圾邮件'
            }
            folder_mapping.update(qq_mapping)
        
        if folder_name in folder_mapping:
            return f"{folder_mapping[folder_name]} ({folder_name})"
        else:
            return folder_name
    
    def search_emails(self, start_date: datetime, end_date: datetime, 
                     folder: str = "INBOX") -> List[str]:
        """
        搜索指定时间范围内的邮件
        
        Args:
            start_date: 开始日期
            end_date: 结束日期
            folder: 邮箱文件夹
            
        Returns:
            邮件ID列表
        """
        if not self.mail:
            raise Exception("未连接到邮箱服务器")
        
        try:
            # 选择文件夹
            self.mail.select(folder)
            
            # 格式化日期查询条件
            old_locale = locale.getlocale(locale.LC_TIME)
            try:
                locale.setlocale(locale.LC_TIME, 'C')
                date_format = '%d-%b-%Y'
                start_date_str = start_date.strftime(date_format)
                end_date_str = (end_date + timedelta(days=1)).strftime(date_format)
                date_query = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
            finally:
                try:
                    if old_locale[0]:
                        locale.setlocale(locale.LC_TIME, old_locale)
                    else:
                        locale.setlocale(locale.LC_TIME, '')
                except:
                    pass
            
            # 搜索邮件
            status, messages = self.mail.search(None, date_query)
            
            if status != 'OK':
                raise Exception(f"搜索邮件失败 - 状态: {status}")
            
            # 获取邮件ID列表
            if messages and messages[0]:
                return messages[0].split()
            else:
                return []
                
        except Exception as e:
            raise Exception(f"搜索邮件失败: {str(e)}")
    
    def fetch_email(self, email_id: str) -> email.message.EmailMessage:
        """
        获取单封邮件
        
        Args:
            email_id: 邮件ID
            
        Returns:
            邮件消息对象
        """
        if not self.mail:
            raise Exception("未连接到邮箱服务器")
        
        try:
            status, data = self.mail.fetch(email_id, '(RFC822)')
            if status != 'OK':
                raise Exception(f"获取邮件失败 - 状态: {status}")
            
            return email.message_from_bytes(data[0][1])
            
        except Exception as e:
            raise Exception(f"获取邮件 {email_id} 失败: {str(e)}")
    
    def fetch_emails_batch(self, start_date: datetime, end_date: datetime,
                          output_file: str, folder: str = "INBOX",
                          download_attachments: bool = False,
                          attachment_folder: str = None,
                          export_format: str = 'csv',
                          progress_callback=None) -> int:
        """
        批量获取并导出邮件
        
        Args:
            start_date: 开始日期
            end_date: 结束日期
            output_file: 输出文件路径
            folder: 邮箱文件夹
            download_attachments: 是否下载附件
            attachment_folder: 附件保存目录
            export_format: 导出格式
            progress_callback: 进度回调函数
            
        Returns:
            导出的邮件数量
        """
        try:
            # 搜索邮件
            if progress_callback:
                progress_callback(0, 0, f"搜索 {folder} 文件夹中的邮件...")
            
            email_ids = self.search_emails(start_date, end_date, folder)
            total_emails = len(email_ids)
            
            if progress_callback:
                progress_callback(0, total_emails, f"找到 {total_emails} 封邮件")
            
            if total_emails == 0:
                if progress_callback:
                    progress_callback(0, 0, "未找到符合条件的邮件")
                return 0
            
            # 处理邮件
            emails_data = []
            batch_size = 50
            processed_count = 0
            
            for batch_start in range(0, total_emails, batch_size):
                batch_end = min(batch_start + batch_size, total_emails)
                batch_ids = email_ids[batch_start:batch_end]
                
                if progress_callback:
                    progress_callback(processed_count, total_emails, 
                                    f"处理第 {batch_start//batch_size + 1} 批邮件")
                
                for i, email_id in enumerate(batch_ids):
                    try:
                        # 获取邮件
                        msg = self.fetch_email(email_id)
                        
                        # 解析邮件
                        email_data = self.parser.parse_complete_email(
                            msg, email_id.decode() if isinstance(email_id, bytes) else email_id,
                            attachment_folder, download_attachments, progress_callback
                        )
                        
                        emails_data.append(email_data)
                        processed_count += 1
                        
                        if progress_callback:
                            subject = email_data.get('subject', '无主题')[:30]
                            progress_callback(processed_count, total_emails, 
                                            f"处理: {subject}{'...' if len(subject) >= 30 else ''}")
                        
                    except Exception as e:
                        if progress_callback:
                            progress_callback(processed_count, total_emails, 
                                            f"跳过邮件 {email_id}: {str(e)[:50]}")
                        continue
                
                # 批次间休息
                if batch_end < total_emails:
                    time.sleep(0.1)
                    
                    # 保持连接活跃
                    if (batch_start // batch_size + 1) % 10 == 0:
                        try:
                            self.mail.noop()
                        except:
                            pass
            
            # 导出邮件
            if progress_callback:
                progress_callback(processed_count, total_emails, "开始导出邮件数据...")
            
            success = self.exporter.export_emails(
                emails_data, output_file, export_format, progress_callback
            )
            
            if success:
                if progress_callback:
                    progress_callback(processed_count, total_emails, 
                                    f"导出完成: {output_file}")
                return len(emails_data)
            else:
                raise Exception("邮件导出失败")
                
        except Exception as e:
            if progress_callback:
                progress_callback(0, 0, f"批量处理失败: {str(e)}")
            raise
    
    def fetch_emails_incremental(self, start_date: datetime, end_date: datetime,
                                output_file: str, folder: str = "INBOX",
                                download_attachments: bool = False,
                                attachment_folder: str = None,
                                export_format: str = 'csv',
                                progress_callback=None,
                                stop_flag=None) -> int:
        """
        增量获取并导出邮件（支持按条写入和停止保存）
        
        Args:
            start_date: 开始日期
            end_date: 结束日期
            output_file: 输出文件路径
            folder: 邮箱文件夹
            download_attachments: 是否下载附件
            attachment_folder: 附件保存目录
            export_format: 导出格式
            progress_callback: 进度回调函数
            stop_flag: 停止标志（可调用对象，返回True表示需要停止）
            
        Returns:
            导出的邮件数量
        """
        incremental_exporter = None
        try:
            # 初始化增量导出器
            incremental_exporter = IncrementalEmailExporter()
            if not incremental_exporter.initialize_export(output_file, export_format):
                raise Exception("初始化增量导出器失败")
            
            # 搜索邮件
            if progress_callback:
                progress_callback(0, 0, f"搜索 {folder} 文件夹中的邮件...")
            
            email_ids = self.search_emails(start_date, end_date, folder)
            total_emails = len(email_ids)
            
            if progress_callback:
                progress_callback(0, total_emails, f"找到 {total_emails} 封邮件")
            
            if total_emails == 0:
                if progress_callback:
                    progress_callback(0, 0, "未找到符合条件的邮件")
                incremental_exporter.finalize_export()
                return 0
            
            # 处理邮件
            processed_count = 0
            batch_size = 10  # 减小批次大小，更频繁地检查停止标志
            
            for batch_start in range(0, total_emails, batch_size):
                # 检查停止标志
                if stop_flag and callable(stop_flag) and stop_flag():
                    if progress_callback:
                        progress_callback(processed_count, total_emails, "用户停止导出，正在保存已处理的邮件...")
                    break
                
                batch_end = min(batch_start + batch_size, total_emails)
                batch_ids = email_ids[batch_start:batch_end]
                
                if progress_callback:
                    progress_callback(processed_count, total_emails, 
                                    f"处理第 {batch_start//batch_size + 1} 批邮件")
                
                for i, email_id in enumerate(batch_ids):
                    # 再次检查停止标志
                    if stop_flag and callable(stop_flag) and stop_flag():
                        if progress_callback:
                            progress_callback(processed_count, total_emails, "用户停止导出，正在保存已处理的邮件...")
                        break
                    
                    try:
                        # 获取邮件
                        msg = self.fetch_email(email_id)
                        
                        # 解析邮件
                        email_data = self.parser.parse_complete_email(
                            msg, email_id.decode() if isinstance(email_id, bytes) else email_id,
                            attachment_folder, download_attachments, progress_callback
                        )
                        
                        # 立即写入文件
                        processed_count += 1
                        if incremental_exporter.add_email(email_data, processed_count):
                            if progress_callback:
                                subject = email_data.get('subject', '无主题')[:30]
                                progress_callback(processed_count, total_emails, 
                                                f"已保存: {subject}{'...' if len(subject) >= 30 else ''}")
                        else:
                            if progress_callback:
                                progress_callback(processed_count, total_emails, 
                                                f"保存邮件 {processed_count} 失败")
                        
                    except Exception as e:
                        if progress_callback:
                            progress_callback(processed_count, total_emails, 
                                            f"跳过邮件 {email_id}: {str(e)[:50]}")
                        continue
                
                # 如果用户停止了，跳出外层循环
                if stop_flag and callable(stop_flag) and stop_flag():
                    break
                
                # 批次间休息
                if batch_end < total_emails:
                    time.sleep(0.1)
                    
                    # 保持连接活跃
                    if (batch_start // batch_size + 1) % 10 == 0:
                        try:
                            self.mail.noop()
                        except:
                            pass
            
            # 完成导出
            if progress_callback:
                if stop_flag and callable(stop_flag) and stop_flag():
                    progress_callback(processed_count, total_emails, "正在保存文件...")
                else:
                    progress_callback(processed_count, total_emails, "正在完成导出...")
            
            success = incremental_exporter.finalize_export()
            
            if success:
                if progress_callback:
                    if stop_flag and callable(stop_flag) and stop_flag():
                        progress_callback(processed_count, total_emails, 
                                        f"已停止并保存: 共导出 {processed_count} 封邮件到 {output_file}")
                    else:
                        progress_callback(processed_count, total_emails, 
                                        f"导出完成: 共导出 {processed_count} 封邮件到 {output_file}")
                return processed_count
            else:
                raise Exception("完成导出失败")
                
        except Exception as e:
            if progress_callback:
                progress_callback(0, 0, f"增量导出失败: {str(e)}")
            raise
        finally:
            # 确保资源被清理
            if incremental_exporter:
                incremental_exporter.cleanup()
    
    def get_provider_info(self) -> Dict[str, Any]:
        """
        获取当前邮箱服务提供商信息
        
        Returns:
            提供商信息字典
        """
        if not self.config:
            return {}
        
        return {
            'name': self.config.name,
            'display_name': self.config.display_name,
            'imap_server': self.config.imap_server,
            'port': self.config.imap_port,
            'auth_type': self.config.auth_type,
            'domain_pattern': self.config.domain_pattern,
            'auth_help': self.providers.get_auth_help(self.config.name)
        }
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()