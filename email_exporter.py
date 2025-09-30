#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件导出器模块
负责将邮件数据导出为各种格式（CSV、JSON等）
"""

import csv
import json
import os
from typing import List, Dict, Any, Optional
from datetime import datetime


class EmailExporter:
    """邮件导出器类"""
    
    def __init__(self):
        self.supported_formats = ['csv', 'json']
    
    def export_to_csv(self, emails: List[Dict[str, Any]], output_file: str, 
                     progress_callback=None) -> bool:
        """导出邮件到CSV文件
        
        Args:
            emails: 邮件数据列表
            output_file: 输出文件路径
            progress_callback: 进度回调函数
            
        Returns:
            导出是否成功
        """
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                # 定义CSV字段
                fieldnames = [
                    '序号', '主题', '发件人', '收件人', '抄送', '密送',
                    '日期', '内容', '附件数量', '附件列表', '邮件ID', '内容类型', '大小'
                ]
                
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                total_emails = len(emails)
                for i, email_data in enumerate(emails, 1):
                    try:
                        # 格式化日期
                        date_str = ""
                        if email_data.get('date'):
                            if isinstance(email_data['date'], datetime):
                                date_str = email_data['date'].strftime('%Y-%m-%d %H:%M:%S')
                            else:
                                date_str = str(email_data['date'])
                        
                        # 格式化附件列表
                        attachment_list = ""
                        if email_data.get('attachment_files'):
                            attachment_list = "; ".join(email_data['attachment_files'])
                        
                        # 写入CSV行
                        writer.writerow({
                            '序号': i,
                            '主题': email_data.get('subject', ''),
                            '发件人': email_data.get('from', ''),
                            '收件人': email_data.get('to', ''),
                            '抄送': email_data.get('cc', ''),
                            '密送': email_data.get('bcc', ''),
                            '日期': date_str,
                            '内容': email_data.get('content', ''),
                            '附件数量': email_data.get('attachment_count', 0),
                            '附件列表': attachment_list,
                            '邮件ID': email_data.get('message_id', ''),
                            '内容类型': email_data.get('content_type', ''),
                            '大小': email_data.get('size', 0)
                        })
                        
                        # 更新进度
                        if progress_callback:
                            progress = int((i / total_emails) * 100)
                            progress_callback(i, total_emails, f"导出邮件 {i}/{total_emails}")
                    
                    except Exception as e:
                        if progress_callback:
                            progress_callback(i, total_emails, f"导出邮件 {i} 失败: {str(e)[:50]}")
                        continue
            
            if progress_callback:
                progress_callback(total_emails, total_emails, f"CSV导出完成: {output_file}")
            
            return True
            
        except Exception as e:
            error_msg = f"CSV导出失败: {str(e)[:100]}"
            if progress_callback:
                progress_callback(0, 0, error_msg)
            return False
    
    def export_to_json(self, emails: List[Dict[str, Any]], output_file: str,
                      progress_callback=None) -> bool:
        """导出邮件到JSON文件
        
        Args:
            emails: 邮件数据列表
            output_file: 输出文件路径
            progress_callback: 进度回调函数
            
        Returns:
            导出是否成功
        """
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 准备JSON数据
            json_data = {
                'export_info': {
                    'export_time': datetime.now().isoformat(),
                    'total_emails': len(emails),
                    'format_version': '1.0'
                },
                'emails': []
            }
            
            total_emails = len(emails)
            for i, email_data in enumerate(emails, 1):
                try:
                    # 处理日期格式
                    email_json = email_data.copy()
                    if email_json.get('date') and isinstance(email_json['date'], datetime):
                        email_json['date'] = email_json['date'].isoformat()
                    
                    # 添加序号
                    email_json['序号'] = i
                    
                    json_data['emails'].append(email_json)
                    
                    # 更新进度
                    if progress_callback:
                        progress = int((i / total_emails) * 100)
                        progress_callback(i, total_emails, f"准备邮件数据 {i}/{total_emails}")
                
                except Exception as e:
                    if progress_callback:
                        progress_callback(i, total_emails, f"处理邮件 {i} 失败: {str(e)[:50]}")
                    continue
            
            # 写入JSON文件
            with open(output_file, 'w', encoding='utf-8') as jsonfile:
                json.dump(json_data, jsonfile, ensure_ascii=False, indent=2)
            
            if progress_callback:
                progress_callback(total_emails, total_emails, f"JSON导出完成: {output_file}")
            
            return True
            
        except Exception as e:
            error_msg = f"JSON导出失败: {str(e)[:100]}"
            if progress_callback:
                progress_callback(0, 0, error_msg)
            return False
    
    def export_emails(self, emails: List[Dict[str, Any]], output_file: str,
                     format_type: str = 'csv', progress_callback=None) -> bool:
        """导出邮件到指定格式
        
        Args:
            emails: 邮件数据列表
            output_file: 输出文件路径
            format_type: 导出格式 ('csv' 或 'json')
            progress_callback: 进度回调函数
            
        Returns:
            导出是否成功
        """
        if not emails:
            error_msg = "没有邮件数据可导出"
            if progress_callback:
                progress_callback(0, 0, error_msg)
            return False
        
        format_type = format_type.lower()
        if format_type not in self.supported_formats:
            error_msg = f"不支持的导出格式: {format_type}。支持的格式: {', '.join(self.supported_formats)}"
            if progress_callback:
                progress_callback(0, 0, error_msg)
            return False
        
        # 根据格式类型调用相应的导出方法
        if format_type == 'csv':
            return self.export_to_csv(emails, output_file, progress_callback)
        elif format_type == 'json':
            return self.export_to_json(emails, output_file, progress_callback)
        
        return False
    
    def get_export_summary(self, emails: List[Dict[str, Any]]) -> Dict[str, Any]:
        """获取导出摘要信息
        
        Args:
            emails: 邮件数据列表
            
        Returns:
            包含摘要信息的字典
        """
        if not emails:
            return {
                'total_emails': 0,
                'total_attachments': 0,
                'date_range': None,
                'senders': [],
                'content_types': []
            }
        
        total_attachments = sum(email.get('attachment_count', 0) for email in emails)
        
        # 统计日期范围
        dates = [email.get('date') for email in emails if email.get('date')]
        date_range = None
        if dates:
            valid_dates = [d for d in dates if isinstance(d, datetime)]
            if valid_dates:
                date_range = {
                    'start': min(valid_dates).isoformat(),
                    'end': max(valid_dates).isoformat()
                }
        
        # 统计发件人
        senders = list(set(email.get('from', '') for email in emails if email.get('from')))
        
        # 统计内容类型
        content_types = list(set(email.get('content_type', '') for email in emails if email.get('content_type')))
        
        return {
            'total_emails': len(emails),
            'total_attachments': total_attachments,
            'date_range': date_range,
            'senders': senders[:10],  # 只显示前10个发件人
            'content_types': content_types
        }
    
    def validate_output_file(self, output_file: str, format_type: str = 'csv') -> bool:
        """验证输出文件路径
        
        Args:
            output_file: 输出文件路径
            format_type: 文件格式类型
            
        Returns:
            路径是否有效
        """
        if not output_file:
            return False
        
        # 检查文件扩展名
        expected_ext = f'.{format_type.lower()}'
        if not output_file.lower().endswith(expected_ext):
            return False
        
        # 检查目录是否可写
        output_dir = os.path.dirname(output_file)
        if output_dir:
            try:
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                return os.access(output_dir, os.W_OK)
            except:
                return False
        
        return True