#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gmail OAuth2认证模块
提供Gmail的OAuth2认证功能
"""

import os
import json
import base64
import webbrowser
from typing import Optional, Dict, Any
from urllib.parse import urlencode, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
import time

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import Flow
    OAUTH_AVAILABLE = True
except ImportError:
    OAUTH_AVAILABLE = False


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """OAuth回调处理器"""
    
    def do_GET(self):
        """处理OAuth回调"""
        if self.path.startswith('/oauth/callback'):
            # 解析授权码
            query_string = self.path.split('?', 1)[1] if '?' in self.path else ''
            params = parse_qs(query_string)
            
            if 'code' in params:
                self.server.auth_code = params['code'][0]
                self.send_response(200)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(b'''
                <html>
                <head><title>\xe6\x8e\x88\xe6\x9d\x83\xe6\x88\x90\xe5\x8a\x9f</title></head>
                <body>
                <h1>\xe6\x8e\x88\xe6\x9d\x83\xe6\x88\x90\xe5\x8a\x9f\xef\xbc\x81</h1>
                <p>\xe6\x82\xa8\xe5\x8f\xaf\xe4\xbb\xa5\xe5\x85\xb3\xe9\x97\xad\xe6\xad\xa4\xe7\xaa\x97\xe5\x8f\xa3\xe5\xb9\xb6\xe8\xbf\x94\xe5\x9b\x9e\xe5\xba\x94\xe7\x94\xa8\xe7\xa8\x8b\xe5\xba\x8f\xe3\x80\x82</p>
                </body>
                </html>
                ''')
            else:
                self.server.auth_error = params.get('error', ['Unknown error'])[0]
                self.send_response(400)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(b'''
                <html>
                <head><title>\xe6\x8e\x88\xe6\x9d\x83\xe5\xa4\xb1\xe8\xb4\xa5</title></head>
                <body>
                <h1>\xe6\x8e\x88\xe6\x9d\x83\xe5\xa4\xb1\xe8\xb4\xa5</h1>
                <p>\xe8\xaf\xb7\xe9\x87\x8d\xe8\xaf\x95\xe6\x88\x96\xe6\xa3\x80\xe6\x9f\xa5\xe9\x85\x8d\xe7\xbd\xae\xe3\x80\x82</p>
                </body>
                </html>
                ''')
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        """禁用日志输出"""
        pass


class GmailOAuth:
    """Gmail OAuth2认证类"""
    
    # Gmail OAuth2配置
    SCOPES = [
        'https://mail.google.com/',  # Gmail IMAP/SMTP访问
        'https://www.googleapis.com/auth/userinfo.email',
        'https://www.googleapis.com/auth/gmail.readonly',
        'openid'
    ]
    REDIRECT_URI = 'http://localhost:8080/oauth/callback'
    
    def __init__(self, client_id: str = None, client_secret: str = None, 
                 credentials_file: str = None, token_file: str = 'gmail_token.json',
                 status_callback=None):
        """
        初始化Gmail OAuth认证
        
        Args:
            client_id: Google OAuth客户端ID
            client_secret: Google OAuth客户端密钥
            credentials_file: 客户端凭据文件路径
            token_file: 访问令牌存储文件路径
            status_callback: 状态回调函数，用于更新授权状态
        """
        if not OAUTH_AVAILABLE:
            raise ImportError("OAuth2依赖库未安装，请运行: pip install google-auth google-auth-oauthlib google-auth-httplib2")
        
        self.client_id = client_id
        self.client_secret = client_secret
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.credentials = None
        self.status_callback = status_callback
        
        # 如果提供了凭据文件，从文件加载配置
        if credentials_file and os.path.exists(credentials_file):
            self._load_credentials_from_file()
        
        # 尝试加载已存在的令牌
        self._load_existing_token()
    
    def _load_credentials_from_file(self):
        """从凭据文件加载OAuth配置"""
        try:
            with open(self.credentials_file, 'r', encoding='utf-8') as f:
                creds_data = json.load(f)
                
            if 'installed' in creds_data:
                client_config = creds_data['installed']
            elif 'web' in creds_data:
                client_config = creds_data['web']
            else:
                raise ValueError("无效的凭据文件格式")
            
            self.client_id = client_config['client_id']
            self.client_secret = client_config['client_secret']
            
        except Exception as e:
            raise ValueError(f"加载凭据文件失败: {str(e)}")
    
    def _validate_token_scopes(self, credentials) -> bool:
        """验证令牌的范围是否与当前需要的范围匹配"""
        if not credentials or not hasattr(credentials, 'scopes'):
            return False
        
        # 获取令牌的范围
        token_scopes = set(credentials.scopes or [])
        required_scopes = set(self.SCOPES)
        
        # 检查是否包含所有必需的范围
        missing_scopes = required_scopes - token_scopes
        if missing_scopes:
            return False
        
        return True
    
    def _load_existing_token(self) -> bool:
        """加载已存在的访问令牌"""
        if os.path.exists(self.token_file):
            try:
                # 首先尝试标准格式（Google OAuth2库格式）
                self.credentials = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)
                
                # 验证令牌有效性和范围
                if self.credentials:
                    if self.credentials.expired and self.credentials.refresh_token:
                        # 如果令牌过期但有刷新令牌，尝试刷新
                        if self._refresh_token():
                            return True
                    elif self.credentials.valid:
                        # 令牌有效，验证范围
                        if self._validate_token_scopes(self.credentials):
                            return True
                        else:
                            return False
                    
            except Exception as e:
                try:
                    # 尝试简化OAuth格式（向后兼容）
                    with open(self.token_file, 'r') as f:
                        token_data = json.load(f)
                    
                    if 'access_token' in token_data:
                        # 简化OAuth格式
                        self.credentials = Credentials(
                            token=token_data.get('access_token'),
                            refresh_token=token_data.get('refresh_token'),
                            token_uri='https://oauth2.googleapis.com/token',
                            client_id=token_data.get('client_id', self.client_id),
                            client_secret=token_data.get('client_secret', self.client_secret),
                            scopes=token_data.get('scopes', self.SCOPES)
                        )
                        
                        # 验证令牌有效性和范围
                        if self.credentials:
                            if self.credentials.expired and self.credentials.refresh_token:
                                if self._refresh_token():
                                    return True
                            elif self.credentials.valid:
                                if self._validate_token_scopes(self.credentials):
                                    return True
                                else:
                                    return False
                        
                except Exception as e2:
                    pass
                    
                return False
        return False
    
    def _refresh_token(self) -> bool:
        """刷新访问令牌"""
        if self.credentials and self.credentials.expired and self.credentials.refresh_token:
            try:
                self.credentials.refresh(Request())
                self._save_token()
                return True
            except Exception as e:
                # 如果刷新失败，删除无效的令牌文件
                if os.path.exists(self.token_file):
                    try:
                        os.remove(self.token_file)
                    except:
                        pass
                return False
        return False
    
    def _save_token(self):
        """保存访问令牌到文件"""
        if self.credentials:
            with open(self.token_file, 'w', encoding='utf-8') as f:
                f.write(self.credentials.to_json())
    
    def _start_callback_server(self) -> tuple:
        """启动OAuth回调服务器"""
        server = HTTPServer(('localhost', 8080), OAuthCallbackHandler)
        server.auth_code = None
        server.auth_error = None
        
        # 在单独线程中运行服务器
        server_thread = threading.Thread(target=server.serve_forever)
        server_thread.daemon = True
        server_thread.start()
        
        return server, server_thread
    
    def authenticate(self, force_reauth: bool = False) -> bool:
        """
        执行OAuth2认证流程
        
        Args:
            force_reauth: 是否强制重新认证
            
        Returns:
            认证是否成功
        """
        if not self.client_id or not self.client_secret:
            raise ValueError("缺少OAuth客户端ID或密钥")
        
        # 如果强制重新认证，清理旧令牌并执行完整流程
        if force_reauth:
            self._cleanup_invalid_token()
            return self._perform_oauth_flow()
        
        # 尝试加载现有令牌
        if self._load_existing_token():
            # 如果令牌有效，确保保存到文件并返回
            if self.credentials.valid:
                # 确保令牌被保存到文件（可能在内部被刷新了）
                self._save_token()
                return True
            # 如果令牌过期但有刷新令牌，尝试刷新
            elif self.credentials.expired and self.credentials.refresh_token:
                if self._refresh_token():
                    return True
        
        # 如果加载失败或刷新失败，清理无效令牌并执行完整的OAuth流程
        self._cleanup_invalid_token()
        return self._perform_oauth_flow()
    
    def _cleanup_invalid_token(self):
        """清理无效的令牌文件"""
        if os.path.exists(self.token_file):
            try:
                os.remove(self.token_file)
            except Exception as e:
                pass
    
    def _perform_oauth_flow(self) -> bool:
        """执行完整的OAuth认证流程"""
        try:
            # 创建OAuth流程
            flow = Flow.from_client_config(
                {
                    "web": {
                        "client_id": self.client_id,
                        "client_secret": self.client_secret,
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": [self.REDIRECT_URI]
                    }
                },
                scopes=self.SCOPES
            )
            flow.redirect_uri = self.REDIRECT_URI
            
            # 启动回调服务器
            server, server_thread = self._start_callback_server()
            
            try:
                # 生成授权URL
                auth_url, _ = flow.authorization_url(
                    access_type='offline',
                    include_granted_scopes='true',
                    prompt='consent'
                )
                
                if self.status_callback:
                    self.status_callback("正在打开浏览器进行授权...")
                
                # 自动打开浏览器
                try:
                    webbrowser.open(auth_url)
                    if self.status_callback:
                        self.status_callback("浏览器已打开，请在浏览器中完成授权...")
                except Exception as e:
                    if self.status_callback:
                        self.status_callback(f"请手动打开浏览器访问: {auth_url}")
                
                # 等待授权码
                timeout = 300  # 5分钟超时
                start_time = time.time()
                
                if self.status_callback:
                    self.status_callback("等待用户在浏览器中完成授权...")
                
                while server.auth_code is None and server.auth_error is None:
                    if time.time() - start_time > timeout:
                        if self.status_callback:
                            self.status_callback("授权超时")
                        raise TimeoutError("授权超时")
                    time.sleep(0.5)
                
                if server.auth_error:
                    if self.status_callback:
                        self.status_callback(f"授权失败: {server.auth_error}")
                    raise Exception(f"授权失败: {server.auth_error}")
                
                if not server.auth_code:
                    if self.status_callback:
                        self.status_callback("未获取到授权码")
                    raise Exception("未获取到授权码")
                
                if self.status_callback:
                    self.status_callback("正在获取访问令牌...")
                
                # 使用授权码获取访问令牌
                try:
                    flow.fetch_token(code=server.auth_code)
                    self.credentials = flow.credentials
                    
                    # 验证获取到的令牌范围
                    if not self._validate_token_scopes(self.credentials):
                        if self.status_callback:
                            self.status_callback("令牌范围可能不完整，但继续使用")
                    
                except Exception as token_error:
                    error_msg = str(token_error)
                    
                    # 处理常见的OAuth错误
                    if "Scope has changed" in error_msg:
                        if self.status_callback:
                            self.status_callback("检测到范围冲突，正在处理...")
                        
                        # 清理可能存在的旧令牌文件
                        self._cleanup_invalid_token()
                        
                        # 提供更友好的错误信息
                        raise Exception("OAuth认证失败: 应用权限范围已更改，请重新进行OAuth授权")
                        
                    elif "invalid_grant" in error_msg or "Bad Request" in error_msg:
                        if self.status_callback:
                            self.status_callback("授权码无效，请重新授权")
                        
                        # 提供更友好的错误信息
                        raise Exception("OAuth认证失败: 授权码无效或已过期，请重新进行OAuth授权")
                        
                    elif "access_denied" in error_msg:
                        if self.status_callback:
                            self.status_callback("用户拒绝了授权")
                        
                        raise Exception("OAuth认证失败: 用户拒绝了授权请求")
                        
                    else:
                        # 对于其他错误，提供通用的错误处理
                        if self.status_callback:
                            self.status_callback("令牌获取失败")
                        
                        raise Exception(f"OAuth认证失败: {error_msg}")
                
                # 保存令牌
                self._save_token()
                
                if self.status_callback:
                    self.status_callback("OAuth授权完成")
                
                return True
                
            finally:
                server.shutdown()
                
        except Exception as e:
            raise Exception(f"OAuth认证失败: {str(e)}")
    
    def get_access_token(self) -> Optional[str]:
        """获取访问令牌"""
        if self.credentials and self.credentials.valid:
            return self.credentials.token
        return None
    
    def get_oauth_string(self, email: str) -> str:
        """
        生成IMAP OAuth认证字符串
        
        Args:
            email: 邮箱地址
            
        Returns:
            OAuth认证字符串（未编码，imaplib会自动进行base64编码）
        """
        # 检查凭据是否存在
        if not self.credentials:
            raise Exception("OAuth凭据不存在，请先进行认证")
        
        # 如果凭据过期，尝试刷新
        if self.credentials.expired:
            if self.credentials.refresh_token:
                if not self._refresh_token():
                    raise Exception("OAuth访问令牌已过期且刷新失败，请重新认证")
            else:
                raise Exception("OAuth访问令牌已过期且无刷新令牌，请重新认证")
        
        # 再次检查凭据有效性
        if not self.credentials.valid:
            raise Exception("OAuth凭据无效，请重新认证")
        
        access_token = self.credentials.token
        # 构建OAuth字符串，格式：user=email\x01auth=Bearer token\x01\x01
        # imaplib会自动进行base64编码，所以这里返回原始字符串
        auth_string = f'user={email}\x01auth=Bearer {access_token}\x01\x01'
        return auth_string
    
    def revoke_token(self):
        """撤销访问令牌"""
        if self.credentials:
            try:
                # 清理本地令牌文件
                if os.path.exists(self.token_file):
                    os.remove(self.token_file)
                self.credentials = None
                return True
            except Exception:
                return False
        return False
    
    def is_authenticated(self) -> bool:
        """检查是否已认证"""
        return self.credentials is not None and self.credentials.valid
    
    def get_user_email(self) -> Optional[str]:
        """获取用户邮箱地址"""
        if not self.is_authenticated():
            return None
        
        try:
            import requests
            # 使用Google API获取用户信息
            headers = {'Authorization': f'Bearer {self.credentials.token}'}
            response = requests.get('https://www.googleapis.com/oauth2/v2/userinfo', headers=headers)
            
            if response.status_code == 200:
                user_info = response.json()
                return user_info.get('email')
        except Exception:
            pass
        
        return None