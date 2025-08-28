#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮箱服务提供商配置模块
支持多种邮箱服务的IMAP连接配置
"""

from dataclasses import dataclass
from typing import Dict, Optional, Tuple
import re
import json
import os
from pathlib import Path


@dataclass
class EmailProviderConfig:
    """邮箱服务提供商配置"""
    name: str  # 提供商名称
    display_name: str  # 显示名称
    imap_server: str  # IMAP服务器地址
    imap_port: int  # IMAP端口
    use_ssl: bool  # 是否使用SSL
    auth_type: str  # 认证类型: 'password', 'oauth2', 'app_password'
    domain_patterns: list  # 邮箱域名模式
    description: str  # 描述信息
    setup_instructions: str  # 设置说明


class EmailProviders:
    """邮箱服务提供商管理器"""
    
    # 默认的邮箱服务提供商配置（作为备用）
    DEFAULT_PROVIDERS = {
        '163': EmailProviderConfig(
            name='163',
            display_name='163邮箱',
            imap_server='imap.163.com',
            imap_port=993,
            use_ssl=True,
            auth_type='password',
            domain_patterns=['163.com', '126.com', 'yeah.net'],
            description='网易163邮箱服务',
            setup_instructions='使用邮箱密码或客户端授权码登录'
        ),
        
        'gmail': EmailProviderConfig(
            name='gmail',
            display_name='Gmail',
            imap_server='imap.gmail.com',
            imap_port=993,
            use_ssl=True,
            auth_type='app_password',
            domain_patterns=['gmail.com', 'googlemail.com'],
            description='Google Gmail邮箱服务',
            setup_instructions='需要开启两步验证并生成应用专用密码'
        ),
        
        'qq': EmailProviderConfig(
            name='qq',
            display_name='QQ邮箱',
            imap_server='imap.qq.com',
            imap_port=993,
            use_ssl=True,
            auth_type='app_password',
            domain_patterns=['qq.com', 'foxmail.com'],
            description='腾讯QQ邮箱服务',
            setup_instructions='需要开启IMAP服务并使用授权码登录'
        ),
        
        'outlook': EmailProviderConfig(
            name='outlook',
            display_name='Outlook/Hotmail',
            imap_server='outlook.office365.com',
            imap_port=993,
            use_ssl=True,
            auth_type='app_password',
            domain_patterns=['outlook.com', 'hotmail.com', 'live.com', 'msn.com'],
            description='微软Outlook邮箱服务',
            setup_instructions='需要开启两步验证并生成应用密码'
        ),
        
        'yahoo': EmailProviderConfig(
            name='yahoo',
            display_name='Yahoo Mail',
            imap_server='imap.mail.yahoo.com',
            imap_port=993,
            use_ssl=True,
            auth_type='app_password',
            domain_patterns=['yahoo.com', 'yahoo.cn', 'ymail.com'],
            description='雅虎邮箱服务',
            setup_instructions='需要开启两步验证并生成应用密码'
        )
    }
    
    # 实际使用的提供商配置（从配置文件加载或使用默认配置）
    PROVIDERS = {}
    
    @classmethod
    def _load_config_file(cls) -> Dict:
        """从配置文件加载提供商配置
        
        Returns:
            配置文件内容的字典，如果加载失败则返回空字典
        """
        config_file = Path(__file__).parent / 'email_providers_config.json'
        
        try:
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"警告：无法加载配置文件 {config_file}: {e}")
            print("将使用默认配置")
        
        return {}
    
    @classmethod
    def _load_providers_from_config(cls) -> Dict[str, EmailProviderConfig]:
        """从配置文件加载提供商配置
        
        Returns:
            提供商配置字典
        """
        config = cls._load_config_file()
        providers = {}
        
        if 'providers' in config:
            for name, provider_data in config['providers'].items():
                try:
                    providers[name] = EmailProviderConfig(
                        name=provider_data['name'],
                        display_name=provider_data['display_name'],
                        imap_server=provider_data['imap_server'],
                        imap_port=provider_data['imap_port'],
                        use_ssl=provider_data['use_ssl'],
                        auth_type=provider_data['auth_type'],
                        domain_patterns=provider_data['domain_patterns'],
                        description=provider_data['description'],
                        setup_instructions=provider_data['setup_instructions']
                    )
                except KeyError as e:
                    print(f"警告：提供商 {name} 配置不完整，缺少字段 {e}")
                    continue
        
        return providers
    
    @classmethod
    def _initialize_providers(cls):
        """初始化提供商配置"""
        if not cls.PROVIDERS:
            # 尝试从配置文件加载
            cls.PROVIDERS = cls._load_providers_from_config()
            
            # 如果配置文件加载失败或为空，使用默认配置
            if not cls.PROVIDERS:
                print("使用默认邮箱提供商配置")
                cls.PROVIDERS = cls.DEFAULT_PROVIDERS.copy()
    
    @classmethod
    def reload_config(cls):
        """重新加载配置文件"""
        cls.PROVIDERS = {}
        cls._initialize_providers()
    
    @classmethod
    def get_provider_by_email(cls, email: str) -> Optional[EmailProviderConfig]:
        """根据邮箱地址自动检测服务提供商
        
        Args:
            email: 邮箱地址
            
        Returns:
            匹配的邮箱服务提供商配置，如果未找到则返回None
        """
        cls._initialize_providers()
        
        if not email or '@' not in email:
            return None
            
        domain = email.split('@')[1].lower()
        
        for provider in cls.PROVIDERS.values():
            if domain in provider.domain_patterns:
                return provider
                
        return None
    
    @classmethod
    def get_provider_by_name(cls, name: str) -> Optional[EmailProviderConfig]:
        """根据提供商名称获取配置
        
        Args:
            name: 提供商名称
            
        Returns:
            邮箱服务提供商配置，如果未找到则返回None
        """
        cls._initialize_providers()
        return cls.PROVIDERS.get(name.lower())
    
    @classmethod
    def get_all_providers(cls) -> Dict[str, EmailProviderConfig]:
        """获取所有支持的邮箱服务提供商
        
        Returns:
            所有邮箱服务提供商配置的字典
        """
        cls._initialize_providers()
        return cls.PROVIDERS.copy()
    
    @classmethod
    def get_provider_names(cls) -> list:
        """获取所有提供商名称列表
        
        Returns:
            提供商名称列表
        """
        cls._initialize_providers()
        return list(cls.PROVIDERS.keys())
    
    @classmethod
    def get_provider_display_names(cls) -> list:
        """获取所有提供商显示名称列表
        
        Returns:
            提供商显示名称列表
        """
        cls._initialize_providers()
        return [provider.display_name for provider in cls.PROVIDERS.values()]
    
    @classmethod
    def validate_email_format(cls, email: str) -> bool:
        """验证邮箱地址格式
        
        Args:
            email: 邮箱地址
            
        Returns:
            邮箱格式是否有效
        """
        if not email:
            return False
            
        # 基本的邮箱格式验证
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))
    
    @classmethod
    def get_connection_params(cls, email: str, provider_name: str = None) -> Optional[Tuple[str, int, bool]]:
        """获取IMAP连接参数
        
        Args:
            email: 邮箱地址
            provider_name: 指定的提供商名称（可选）
            
        Returns:
            (IMAP服务器, 端口, 是否使用SSL) 或 None
        """
        provider = None
        
        if provider_name:
            provider = cls.get_provider_by_name(provider_name)
        else:
            provider = cls.get_provider_by_email(email)
            
        if provider:
            return provider.imap_server, provider.imap_port, provider.use_ssl
            
        return None
    
    @classmethod
    def get_auth_instructions(cls, email: str, provider_name: str = None) -> Optional[str]:
        """获取认证设置说明
        
        Args:
            email: 邮箱地址
            provider_name: 指定的提供商名称（可选）
            
        Returns:
            认证设置说明文本
        """
        provider = None
        
        if provider_name:
            provider = cls.get_provider_by_name(provider_name)
        else:
            provider = cls.get_provider_by_email(email)
            
        if provider:
            return provider.setup_instructions
            
        return None


def detect_email_provider(email: str) -> Optional[str]:
    """检测邮箱服务提供商（便捷函数）
    
    Args:
        email: 邮箱地址
        
    Returns:
        提供商名称，如果未检测到则返回None
    """
    provider = EmailProviders.get_provider_by_email(email)
    return provider.name if provider else None


def get_imap_config(email: str, provider_name: str = None) -> Optional[dict]:
    """获取IMAP配置信息（便捷函数）
    
    Args:
        email: 邮箱地址
        provider_name: 指定的提供商名称（可选）
        
    Returns:
        包含IMAP配置信息的字典
    """
    provider = None
    
    if provider_name:
        provider = EmailProviders.get_provider_by_name(provider_name)
    else:
        provider = EmailProviders.get_provider_by_email(email)
        
    if provider:
        return {
            'server': provider.imap_server,
            'port': provider.imap_port,
            'use_ssl': provider.use_ssl,
            'auth_type': provider.auth_type,
            'provider_name': provider.name,
            'display_name': provider.display_name,
            'instructions': provider.setup_instructions
        }
        
    return None