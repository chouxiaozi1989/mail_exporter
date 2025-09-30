#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增量邮件导出器模块
支持按条写入和停止时保存功能
"""

import csv
import json
import os
from typing import List, Dict, Any, Optional
from datetime import datetime
import threading


class IncrementalEmailExporter:
    """增量邮件导出器类"""
    
    def __init__(self):
        self.supported_formats = ['csv', 'json']
        self.current_file = None
        self.csv_writer = None
        self.json_data = None
        self.output_file = None
        self.format_type = None
        self.lock = threading.Lock()
        self.is_initialized = False
        
    def initialize_export(self, output_file: str, format_type: str = 'csv') -> bool:
        """初始化导出文件
        
        Args:
            output_file: 输出文件路径
            format_type: 导出格式 ('csv' 或 'json')
            
        Returns:
            初始化是否成功
        """
        try:
            with self.lock:
                self.output_file = output_file
                self.format_type = format_type.lower()
                
                if self.format_type not in self.supported_formats:
                    raise ValueError(f"不支持的导出格式: {format_type}")
                
                # 确保输出目录存在
                output_dir = os.path.dirname(output_file)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                if self.format_type == 'csv':
                    self._initialize_csv()
                elif self.format_type == 'json':
                    self._initialize_json()
                
                self.is_initialized = True
                return True
                
        except Exception as e:
            return False
    
    def _initialize_csv(self):
        """初始化CSV文件"""
        self.current_file = open(self.output_file, 'w', newline='', encoding='utf-8-sig')
        
        # 定义CSV字段
        fieldnames = [
            '序号', '主题', '发件人', '收件人', '抄送', '密送',
            '日期', '内容', '附件数量', '附件列表', '邮件ID', '内容类型', '大小'
        ]
        
        self.csv_writer = csv.DictWriter(self.current_file, fieldnames=fieldnames)
        self.csv_writer.writeheader()
        self.current_file.flush()
    
    def _initialize_json(self):
        """初始化JSON文件"""
        self.json_data = {
            'export_info': {
                'export_time': datetime.now().isoformat(),
                'total_emails': 0,
                'format_version': '1.0'
            },
            'emails': []
        }
    
    def add_email(self, email_data: Dict[str, Any], sequence_number: int) -> bool:
        """添加一封邮件到导出文件
        
        Args:
            email_data: 邮件数据
            sequence_number: 序号
            
        Returns:
            添加是否成功
        """
        if not self.is_initialized:
            return False
            
        try:
            with self.lock:
                if self.format_type == 'csv':
                    return self._add_email_to_csv(email_data, sequence_number)
                elif self.format_type == 'json':
                    return self._add_email_to_json(email_data, sequence_number)
                    
        except Exception as e:
            return False
            
        return False
    
    def _add_email_to_csv(self, email_data: Dict[str, Any], sequence_number: int) -> bool:
        """添加邮件到CSV文件"""
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
            self.csv_writer.writerow({
                '序号': sequence_number,
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
            
            # 立即刷新到磁盘
            self.current_file.flush()
            return True
            
        except Exception as e:
            return False
    
    def _add_email_to_json(self, email_data: Dict[str, Any], sequence_number: int) -> bool:
        """添加邮件到JSON数据"""
        try:
            # 处理日期格式
            email_json = email_data.copy()
            if email_json.get('date') and isinstance(email_json['date'], datetime):
                email_json['date'] = email_json['date'].isoformat()
            
            # 添加序号
            email_json['序号'] = sequence_number
            
            # 添加到JSON数据
            self.json_data['emails'].append(email_json)
            self.json_data['export_info']['total_emails'] = len(self.json_data['emails'])
            
            return True
            
        except Exception as e:
            return False
    
    def finalize_export(self) -> bool:
        """完成导出，保存文件
        
        Returns:
            完成是否成功
        """
        if not self.is_initialized:
            return False
            
        try:
            with self.lock:
                if self.format_type == 'csv':
                    return self._finalize_csv()
                elif self.format_type == 'json':
                    return self._finalize_json()
                    
        except Exception as e:
            return False
        finally:
            self.is_initialized = False
            
        return False
    
    def _finalize_csv(self) -> bool:
        """完成CSV导出"""
        try:
            if self.current_file:
                self.current_file.flush()
                self.current_file.close()
                self.current_file = None
                self.csv_writer = None
            return True
            
        except Exception as e:
            return False
    
    def _finalize_json(self) -> bool:
        """完成JSON导出"""
        try:
            # 写入JSON文件
            with open(self.output_file, 'w', encoding='utf-8') as jsonfile:
                json.dump(self.json_data, jsonfile, ensure_ascii=False, indent=2)
            
            self.json_data = None
            return True
            
        except Exception as e:
            return False
    
    def cleanup(self):
        """清理资源"""
        try:
            with self.lock:
                if self.current_file:
                    self.current_file.close()
                    self.current_file = None
                self.csv_writer = None
                self.json_data = None
                self.is_initialized = False
        except:
            pass
    
    def __del__(self):
        """析构函数，确保资源被释放"""
        self.cleanup()