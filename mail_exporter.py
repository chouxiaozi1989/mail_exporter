#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件导出工具 - 重构版本
支持多种邮箱服务提供商（163、Gmail、QQ等）
"""

import argparse
import getpass
import sys
import os
from datetime import datetime
from email_providers import EmailProviders
from mail_client import MailClient
from oauth_gmail import GmailOAuth

def get_supported_providers():
    """获取支持的邮箱服务提供商列表
    
    Returns:
        支持的提供商列表
    """
    providers = EmailProviders()
    return providers.get_all_providers()

def get_mail_folders(username, password, provider=None, proxy_config=None, oauth_config=None):
    """获取邮箱文件夹列表"""
    try:
        # 使用MailClient获取文件夹列表
        if provider:
            client = MailClient(provider_name=provider, proxy_config=proxy_config, oauth_config=oauth_config)
        else:
            client = MailClient(email_address=username, proxy_config=proxy_config, oauth_config=oauth_config)
        client.connect(username, password)
        folders = client.get_folders()
        client.disconnect()
        return folders
    except Exception as e:
        raise Exception(f"获取文件夹列表失败: {str(e)}")

def fetch_emails(username, password, start_date, end_date, output_file, folder="INBOX", 
                progress_callback=None, download_attachments=False, attachment_folder=None, provider=None, proxy_config=None):
    """从邮箱获取邮件并导出"""
    try:
        # 使用MailClient进行邮件获取和导出
        if provider:
            client = MailClient(provider_name=provider, proxy_config=proxy_config)
        else:
            client = MailClient(email_address=username, proxy_config=proxy_config)
        client.connect(username, password)
        
        email_count = client.fetch_emails_batch(
            start_date=start_date,
            end_date=end_date,
            output_file=output_file,
            folder=folder,
            download_attachments=download_attachments,
            attachment_folder=attachment_folder,
            progress_callback=progress_callback
        )
        
        client.disconnect()
        return email_count
        
    except Exception as e:
        raise Exception(f"邮件导出失败: {str(e)}")

def fetch_emails_incremental(username, password, start_date, end_date, output_file, folder="INBOX", 
                            progress_callback=None, download_attachments=False, attachment_folder=None, 
                            provider=None, stop_flag=None, proxy_config=None, oauth_config=None, email_count_limit=0):
    """从邮箱增量获取邮件并导出（支持按条写入和停止保存）"""
    try:
        # 使用MailClient进行邮件获取和导出
        if provider:
            client = MailClient(provider_name=provider, proxy_config=proxy_config, oauth_config=oauth_config)
        else:
            client = MailClient(email_address=username, proxy_config=proxy_config, oauth_config=oauth_config)
        client.connect(username, password)
        
        email_count = client.fetch_emails_incremental(
            start_date=start_date,
            end_date=end_date,
            output_file=output_file,
            folder=folder,
            download_attachments=download_attachments,
            attachment_folder=attachment_folder,
            progress_callback=progress_callback,
            stop_flag=stop_flag,
            email_count_limit=email_count_limit
        )
        
        client.disconnect()
        return email_count
        
    except Exception as e:
        raise Exception(f"增量邮件导出失败: {str(e)}")

def test_oauth_auth(oauth_config):
    """
    测试OAuth授权
    
    Args:
        oauth_config: OAuth配置字典
        
    Returns:
        bool: 授权是否成功
    """
    try:
        gmail_oauth = GmailOAuth(
            client_id=oauth_config.get('client_id'),
            client_secret=oauth_config.get('client_secret'),
            credentials_file=oauth_config.get('credentials_file'),
            token_file=oauth_config.get('token_file', 'gmail_token.json')
        )
        
        # 执行OAuth认证，强制重新认证以确保获取最新授权
        return gmail_oauth.authenticate(force_reauth=True)
        
    except Exception as e:
        print(f"OAuth授权测试失败: {str(e)}")
        return False

# 邮件处理函数已移至模块化文件中

# decode_subject 函数已移至模块化文件中

# get_mail_from 函数已移至模块化文件中


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='从邮箱导出邮件到CSV/JSON文件，支持多种邮箱服务提供商')
    parser.add_argument('-u', '--username', help='邮箱用户名')
    parser.add_argument('-p', '--password', help='邮箱密码或授权码（如果不提供将提示输入）')
    parser.add_argument('-s', '--start-date', help='开始日期 (YYYY-MM-DD)')
    parser.add_argument('-e', '--end-date', help='结束日期 (YYYY-MM-DD)')
    parser.add_argument('-o', '--output', help='输出文件路径（支持.csv和.json格式）')
    parser.add_argument('-f', '--folder', default='INBOX', help='邮箱文件夹名称，默认为INBOX')
    parser.add_argument('-a', '--attachments', action='store_true', help='下载邮件附件')
    parser.add_argument('--attachment-folder', help='附件保存文件夹路径（默认为输出文件同目录下的attachments文件夹）')
    parser.add_argument('--provider', help='邮箱服务提供商（163、gmail、qq、outlook、yahoo等，不指定则自动检测）')
    parser.add_argument('--list-providers', action='store_true', help='列出支持的邮箱服务提供商')
    parser.add_argument('--list-folders', action='store_true', help='列出邮箱中的所有文件夹')
    
    args = parser.parse_args()
    
    # 列出支持的提供商
    if args.list_providers:
        print("支持的邮箱服务提供商:")
        providers = get_supported_providers()
        for name, config in providers.items():
            print(f"  {name}: {config.display_name}")
            print(f"    IMAP服务器: {config.imap_server}:{config.imap_port}")
            print(f"    认证类型: {config.auth_type}")
            if config.domain_patterns:
                print(f"    域名模式: {', '.join(config.domain_patterns)}")
            print()
        sys.exit(0)
    
    # 列出文件夹
    if args.list_folders:
        if not args.username:
            parser.error("列出文件夹需要提供邮箱用户名 (-u/--username)")
        # 如果没有提供密码，提示用户输入
        if not args.password:
            provider_info = ""
            if args.provider:
                provider_info = f"({args.provider}) "
            args.password = getpass.getpass(f'请输入邮箱 {provider_info}密码或授权码: ')
        try:
            print("正在获取邮箱文件夹列表...")
            folders = get_mail_folders(args.username, args.password, args.provider)
            print("\n可用的邮箱文件夹:")
            for folder_name, display_name in folders:
                print(f"  {folder_name}: {display_name}")
        except Exception as e:
            print(f"获取文件夹列表失败: {e}")
            sys.exit(1)
        sys.exit(0)
    
    # 检查必需参数
    if not args.username:
        parser.error("必须提供邮箱用户名 (-u/--username)")
    
    if not args.output:
        parser.error("必须提供输出文件路径 (-o/--output)")
    
    # 如果没有提供密码，提示用户输入
    if not args.password:
        provider_info = ""
        if args.provider:
            provider_info = f"({args.provider}) "
        args.password = getpass.getpass(f'请输入邮箱 {provider_info}密码或授权码: ')
    
    # 解析日期
    start_date = None
    end_date = None
    
    if args.start_date:
        try:
            start_date = datetime.strptime(args.start_date, '%Y-%m-%d')
        except ValueError:
            print("错误: 开始日期格式不正确，请使用 YYYY-MM-DD 格式")
            sys.exit(1)
    
    if args.end_date:
        try:
            end_date = datetime.strptime(args.end_date, '%Y-%m-%d')
        except ValueError:
            print("错误: 结束日期格式不正确，请使用 YYYY-MM-DD 格式")
            sys.exit(1)
    
    # 验证日期范围
    if start_date and end_date and start_date > end_date:
        print("错误: 开始日期不能晚于结束日期")
        sys.exit(1)
    
    # 检测邮箱服务提供商
    detected_provider = None
    if not args.provider:
        try:
            providers = EmailProviders()
            config = providers.get_provider_by_email(args.username)
            if config:
                detected_provider = config.name
                print(f"自动检测到邮箱服务提供商: {config.display_name}")
        except Exception:
            pass
    
    # 显示配置信息
    print(f"\n=== 邮件导出配置 ===")
    print(f"邮箱用户名: {args.username}")
    print(f"邮箱服务商: {args.provider or detected_provider or '自动检测'}")
    print(f"开始日期: {args.start_date if args.start_date else '不限制'}")
    print(f"结束日期: {args.end_date if args.end_date else '不限制'}")
    print(f"输出文件: {args.output}")
    print(f"邮箱文件夹: {args.folder}")
    print(f"下载附件: {'是' if args.attachments else '否'}")
    if args.attachments:
        attachment_folder = args.attachment_folder
        if not attachment_folder:
            # 使用输出文件同目录下的attachments文件夹
            output_dir = os.path.dirname(os.path.abspath(args.output))
            attachment_folder = os.path.join(output_dir, 'attachments')
        print(f"附件保存路径: {attachment_folder}")
    print()
    
    try:
        # 调用邮件获取函数
        email_count = fetch_emails(
            username=args.username,
            password=args.password,
            start_date=start_date,
            end_date=end_date,
            output_file=args.output,
            folder=args.folder,
            download_attachments=args.attachments,
            attachment_folder=args.attachment_folder,
            provider=args.provider or detected_provider
        )
        
        print(f"\n=== 导出完成 ===")
        print(f"成功导出 {email_count} 封邮件")
        
    except KeyboardInterrupt:
        print("\n用户中断操作")
        sys.exit(1)
    except Exception as e:
        print(f"\n导出失败: {e}")
        sys.exit(1)