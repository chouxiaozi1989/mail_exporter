#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件解析器模块
负责解析邮件内容、附件和元数据
"""

import email
import email.header
import email.utils
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import base64
import quopri
import re
import os
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any, Union


class EmailParser:
    """邮件解析器类"""
    
    def __init__(self):
        self.attachment_count = 0
        self.attachment_files = []
    
    def decode_imap_utf7(self, text: str) -> str:
        """解码IMAP UTF-7编码的文本
        
        Args:
            text: 需要解码的文本
            
        Returns:
            解码后的文本
        """
        try:
            # 尝试使用标准库解码
            import imaplib
            decoded = text.encode('ascii').decode('imap4-utf-7')
            return decoded
        except:
            # 如果标准库解码失败，使用手动解码
            return self._manual_decode_imap_utf7(text)
    
    def _manual_decode_imap_utf7(self, text: str) -> str:
        """手动解码IMAP UTF-7编码
        
        Args:
            text: 需要解码的文本
            
        Returns:
            解码后的文本
        """
        try:
            # 简单的IMAP UTF-7解码实现
            result = []
            i = 0
            while i < len(text):
                if text[i] == '&':
                    if i + 1 < len(text) and text[i + 1] == '-':
                        result.append('&')
                        i += 2
                    else:
                        # 查找结束的'-'
                        end = text.find('-', i + 1)
                        if end != -1:
                            encoded = text[i + 1:end]
                            try:
                                # 添加填充并解码
                                padding = 4 - len(encoded) % 4
                                if padding != 4:
                                    encoded += '=' * padding
                                import base64
                                decoded_bytes = base64.b64decode(encoded)
                                decoded_text = decoded_bytes.decode('utf-16be')
                                result.append(decoded_text)
                            except:
                                result.append(text[i:end + 1])
                            i = end + 1
                        else:
                            result.append(text[i])
                            i += 1
                else:
                    result.append(text[i])
                    i += 1
            return ''.join(result)
        except Exception as e:
            return text
    
    def decode_subject(self, subject: str) -> str:
        """解码邮件主题
        
        Args:
            subject: 原始邮件主题
            
        Returns:
            解码后的主题字符串
        """
        if subject:
            try:
                # 尝试解码 Subject
                decoded_parts = email.header.decode_header(subject)
                subject_str = ""
                for part, charset in decoded_parts:
                    if isinstance(part, bytes):
                        # 优先使用返回的编码集，如果不可用则使用utf-8
                        if charset:
                            try:
                                subject_str += part.decode(charset, errors='ignore')
                            except (LookupError, UnicodeDecodeError):
                                subject_str += part.decode('utf-8', errors='ignore')
                        else:
                            subject_str += part.decode('utf-8', errors='ignore')
                    else:
                        subject_str += part
                return subject_str
            except Exception as e:
                return subject
        else:
            return "[无主题]"
    
    def get_mail_from(self, from_: str) -> str:
        """解析邮件发件人信息
        
        Args:
            from_: 原始发件人头信息
            
        Returns:
            发件人邮箱地址
        """
        if from_:
            try:
                # 解析发件人信息
                decoded_parts = email.header.decode_header(from_)
                from_str = ""
                for part, charset in decoded_parts:
                    if isinstance(part, bytes):
                        # 优先使用返回的编码集，如果不可用则使用utf-8
                        if charset:
                            try:
                                from_str += part.decode(charset, errors='ignore')
                            except (LookupError, UnicodeDecodeError):
                                from_str += part.decode('utf-8', errors='ignore')
                        else:
                            from_str += part.decode('utf-8', errors='ignore')
                    else:
                        from_str += part
                
                # 提取邮箱地址
                # 格式可能是 "姓名 <email@example.com>" 或纯邮箱地址
                email_match = re.search(r'<([^<>]+)>', from_str)
                if email_match:
                    # 只返回邮箱地址部分
                    return email_match.group(1).strip()
                else:
                    # 如果没有尖括号格式，检查是否为纯邮箱地址
                    # 更宽松的邮箱正则表达式，支持更多有效格式
                    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                    if re.match(email_pattern, from_str.strip()):
                        return from_str.strip()
                    else:
                        # 尝试从字符串中提取邮箱地址
                        email_in_text = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', from_str)
                        if email_in_text:
                            return email_in_text.group(0)
                        else:
                            return from_str
            except Exception as e:
                return from_
        else:
            return "[未知发件人]"
    
    def parse_email_content(self, msg) -> str:
        """解析邮件正文内容
        
        Args:
            msg: 邮件消息对象
            
        Returns:
            解析后的邮件正文
        """
        content = ""
        
        try:
            if msg.is_multipart():
                # 多部分邮件
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition", ""))
                    
                    # 跳过附件
                    if "attachment" in content_disposition:
                        continue
                    
                    # 处理文本内容
                    if content_type == "text/plain":
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                charset = part.get_content_charset() or 'utf-8'
                                try:
                                    text = payload.decode(charset, errors='ignore')
                                except (LookupError, UnicodeDecodeError):
                                    text = payload.decode('utf-8', errors='ignore')
                                content += text + "\n"
                        except Exception as e:
                            pass
                    
                    elif content_type == "text/html":
                        try:
                            payload = part.get_payload(decode=True)
                            if payload:
                                charset = part.get_content_charset() or 'utf-8'
                                try:
                                    html_text = payload.decode(charset, errors='ignore')
                                except (LookupError, UnicodeDecodeError):
                                    html_text = payload.decode('utf-8', errors='ignore')
                                
                                # 简单的HTML标签清理
                                import re
                                clean_text = re.sub(r'<[^>]+>', '', html_text)
                                clean_text = re.sub(r'\s+', ' ', clean_text).strip()
                                if clean_text and not content.strip():
                                    content = clean_text + "\n"
                        except Exception as e:
                            pass
            else:
                # 单部分邮件
                try:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        charset = msg.get_content_charset() or 'utf-8'
                        try:
                            content = payload.decode(charset, errors='ignore')
                        except (LookupError, UnicodeDecodeError):
                            content = payload.decode('utf-8', errors='ignore')
                except Exception as e:
                    pass
        
        except Exception as e:
            pass
        
        return content.strip() if content else "[无内容]"
    
    def process_attachments(self, msg, 
                          attachment_folder: str, email_id: str,
                          progress_callback=None) -> Tuple[int, List[str]]:
        """处理邮件附件
        
        Args:
            msg: 邮件消息对象
            attachment_folder: 附件保存根目录
            email_id: 邮件ID
            progress_callback: 进度回调函数
            
        Returns:
            (附件数量, 附件文件列表)
        """
        attachment_count = 0
        attachment_files = []
        
        if not msg.is_multipart():
            return attachment_count, attachment_files
        
        # 为每封邮件创建单独的附件文件夹
        email_attachment_folder = os.path.join(attachment_folder, f"email_{email_id}")
        
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # 检查是否为附件
            if "attachment" in content_disposition or part.get_filename():
                try:
                    filename = part.get_filename()
                    if filename:
                        # 解码文件名
                        try:
                            decoded_parts = email.header.decode_header(filename)
                            filename_str = ""
                            for part_data, charset in decoded_parts:
                                if isinstance(part_data, bytes):
                                    if charset:
                                        try:
                                            filename_str += part_data.decode(charset, errors='ignore')
                                        except (LookupError, UnicodeDecodeError):
                                            filename_str += part_data.decode('utf-8', errors='ignore')
                                    else:
                                        filename_str += part_data.decode('utf-8', errors='ignore')
                                else:
                                    filename_str += part_data
                            filename = filename_str
                        except:
                            # 如果解码失败，使用原始文件名
                            filename = filename_str
                        
                        # 创建邮件附件文件夹
                        if not os.path.exists(email_attachment_folder):
                            os.makedirs(email_attachment_folder)
                        
                        # 生成安全的文件名
                        safe_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                        if not safe_filename:
                            safe_filename = f"attachment_{attachment_count + 1}"
                        
                        file_path = os.path.join(email_attachment_folder, safe_filename)
                        
                        # 如果文件已存在，添加数字后缀
                        counter = 1
                        original_path = file_path
                        while os.path.exists(file_path):
                            name, ext = os.path.splitext(original_path)
                            file_path = f"{name}_{counter}{ext}"
                            counter += 1
                        
                        # 保存附件
                        payload = part.get_payload(decode=True)
                        if payload:
                            with open(file_path, 'wb') as f:
                                f.write(payload)
                            
                            attachment_count += 1
                            attachment_files.append(os.path.relpath(file_path, attachment_folder))
                            
                            if progress_callback:
                                progress_callback(0, 0, f"保存附件: {safe_filename}")
                            
                except Exception as e:
                    if progress_callback:
                        progress_callback(0, 0, f"保存附件失败: {filename} - {str(e)[:50]}")
        
        return attachment_count, attachment_files
    
    def parse_email_date(self, date_str: str) -> Optional[datetime]:
        """解析邮件日期
        
        Args:
            date_str: 邮件日期字符串
            
        Returns:
            解析后的日期时间对象
        """
        if not date_str:
            return None
            
        try:
            # 尝试使用email.utils解析日期
            import email.utils
            parsed_date = email.utils.parsedate_tz(date_str)
            if parsed_date:
                timestamp = email.utils.mktime_tz(parsed_date)
                return datetime.fromtimestamp(timestamp)
        except Exception as e:
            return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
        return None
    
    def extract_email_metadata(self, msg) -> Dict[str, Any]:
        """提取邮件元数据
        
        Args:
            msg: 邮件消息对象
            
        Returns:
            包含邮件元数据的字典
        """
        metadata = {
            'subject': self.decode_subject(msg.get('Subject', '')),
            'from': self.get_mail_from(msg.get('From', '')),
            'to': msg.get('To', ''),
            'cc': msg.get('Cc', ''),
            'bcc': msg.get('Bcc', ''),
            'date': self.parse_email_date(msg.get('Date', '')),
            'message_id': msg.get('Message-ID', ''),
            'content_type': msg.get_content_type(),
            'size': len(str(msg))
        }
        
        return metadata
    
    def parse_complete_email(self, msg, 
                           email_id: str = None,
                           attachment_folder: str = None,
                           download_attachments: bool = False,
                           progress_callback=None) -> Dict[str, Any]:
        """完整解析邮件
        
        Args:
            msg: 邮件消息对象
            email_id: 邮件ID
            attachment_folder: 附件保存目录
            download_attachments: 是否下载附件
            progress_callback: 进度回调函数
            
        Returns:
            包含完整邮件信息的字典
        """
        # 提取元数据
        email_data = self.extract_email_metadata(msg)
        
        # 解析邮件内容
        email_data['content'] = self.parse_email_content(msg)
        
        # 处理附件
        if download_attachments and attachment_folder and email_id:
            attachment_count, attachment_files = self.process_attachments(
                msg, attachment_folder, email_id, progress_callback
            )
            email_data['attachment_count'] = attachment_count
            email_data['attachment_files'] = attachment_files
        else:
            email_data['attachment_count'] = 0
            email_data['attachment_files'] = []
        
        return email_data