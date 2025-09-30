#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
支持代理的IMAP连接模块
"""

import imaplib
import socket
import ssl
import urllib.request
import urllib.parse
import http.client
try:
    import socks
    SOCKS_AVAILABLE = True
except ImportError:
    SOCKS_AVAILABLE = False


class ProxyIMAP4(imaplib.IMAP4):
    """支持代理的IMAP4类"""
    
    def __init__(self, host='', port=imaplib.IMAP4_PORT, proxy_host=None, proxy_port=None, proxy_type='http', proxy_username=None, proxy_password=None):
        self.proxy_host = proxy_host
        self.proxy_port = proxy_port
        self.proxy_type = proxy_type.lower() if proxy_type else 'http'
        self.proxy_username = proxy_username
        self.proxy_password = proxy_password
        super().__init__(host, port)
    
    def _create_socket(self, timeout=socket._GLOBAL_DEFAULT_TIMEOUT):
        """创建支持代理的socket连接"""
        if self.proxy_host and self.proxy_port:
            try:
                return self._create_proxy_socket(timeout)
            except Exception:
                # 如果代理连接失败，回退到直连
                pass
        
        # 直连
        return socket.create_connection((self.host, self.port), timeout)
    
    def _create_proxy_socket(self, timeout):
        """创建代理socket连接"""
        if self.proxy_type in ['socks5', 'socks4'] and SOCKS_AVAILABLE:
            return self._create_socks_socket(timeout)
        elif self.proxy_type == 'http':
            return self._create_http_proxy_socket(timeout)
        else:
            raise Exception(f"不支持的代理类型: {self.proxy_type}")
    
    def _create_socks_socket(self, timeout):
        """创建SOCKS代理socket连接"""
        if not SOCKS_AVAILABLE:
            raise Exception("SOCKS代理需要安装PySocks库: pip install PySocks")
        
        # 创建SOCKS socket
        sock = socks.socksocket()
        sock.settimeout(timeout)
        
        # 设置代理
        if self.proxy_type == 'socks5':
            proxy_type = socks.SOCKS5
        elif self.proxy_type == 'socks4':
            proxy_type = socks.SOCKS4
        else:
            raise Exception(f"不支持的SOCKS代理类型: {self.proxy_type}")
        
        sock.set_proxy(proxy_type, self.proxy_host, self.proxy_port, 
                      username=self.proxy_username, password=self.proxy_password)
        
        # 连接到目标服务器
        sock.connect((self.host, self.port))
        return sock
    
    def _create_http_proxy_socket(self, timeout):
        """创建HTTP代理socket连接"""
        # 创建到代理服务器的连接
        proxy_sock = socket.create_connection((self.proxy_host, self.proxy_port), timeout)
        
        # 发送CONNECT请求
        connect_req = f"CONNECT {self.host}:{self.port} HTTP/1.1\r\n"
        connect_req += f"Host: {self.host}:{self.port}\r\n"
        
        # 添加代理认证
        if self.proxy_username and self.proxy_password:
            import base64
            auth_string = f"{self.proxy_username}:{self.proxy_password}"
            auth_bytes = base64.b64encode(auth_string.encode('utf-8')).decode('ascii')
            connect_req += f"Proxy-Authorization: Basic {auth_bytes}\r\n"
        
        connect_req += "\r\n"
        
        proxy_sock.send(connect_req.encode('utf-8'))
        
        # 读取响应
        response = proxy_sock.recv(4096).decode('utf-8')
        if not response.startswith('HTTP/1.1 200') and not response.startswith('HTTP/1.0 200'):
            proxy_sock.close()
            raise Exception(f"代理连接失败: {response.split()[1] if len(response.split()) > 1 else 'Unknown error'}")
        
        return proxy_sock


class ProxyIMAP4_SSL(ProxyIMAP4):
    """支持代理的IMAP4_SSL类"""
    
    def __init__(self, host='', port=imaplib.IMAP4_SSL_PORT, keyfile=None, certfile=None, 
                 ssl_context=None, proxy_host=None, proxy_port=None, proxy_type='http', proxy_username=None, proxy_password=None):
        self.keyfile = keyfile
        self.certfile = certfile
        self.ssl_context = ssl_context
        # 调用ProxyIMAP4的初始化，而不是IMAP4_SSL
        ProxyIMAP4.__init__(self, host, port, proxy_host, proxy_port, proxy_type, proxy_username, proxy_password)
    
    def _create_socket(self, timeout=socket._GLOBAL_DEFAULT_TIMEOUT):
        """创建支持代理的SSL socket连接"""
        # 创建基础socket连接（支持代理）
        sock = super()._create_socket(timeout)
        
        # 包装为SSL socket
        if self.ssl_context:
            sock = self.ssl_context.wrap_socket(sock, server_hostname=self.host)
        else:
            # 使用现代的SSL上下文而不是已弃用的ssl.wrap_socket
            context = ssl.create_default_context()
            
            # 针对Gmail的SSL配置优化
            context.check_hostname = True
            context.verify_mode = ssl.CERT_REQUIRED
            
            # 设置更宽松的SSL选项以避免EOF错误
            context.options |= ssl.OP_NO_SSLv2
            context.options |= ssl.OP_NO_SSLv3
            context.options |= ssl.OP_SINGLE_DH_USE
            context.options |= ssl.OP_SINGLE_ECDH_USE
            
            # 允许不安全的重新协商（某些服务器需要）
            try:
                context.options |= ssl.OP_LEGACY_SERVER_CONNECT
            except AttributeError:
                pass  # 旧版本Python可能没有这个选项
            
            if self.keyfile and self.certfile:
                context.load_cert_chain(self.certfile, self.keyfile)
            
            # 使用更robust的SSL包装，添加超时和错误处理
            try:
                sock = context.wrap_socket(sock, server_hostname=self.host)
            except ssl.SSLError as e:
                # 如果SSL握手失败，尝试降级处理
                sock.close()
                raise Exception(f"SSL连接失败: {str(e)}。请检查网络连接或尝试禁用代理。")
        
        return sock


def get_system_proxy():
    """获取系统代理设置"""
    try:
        # 尝试获取系统HTTP代理
        proxy_handler = urllib.request.getproxies()
        
        if 'http' in proxy_handler:
            proxy_url = proxy_handler['http']
            parsed = urllib.parse.urlparse(proxy_url)
            return {
                'host': parsed.hostname,
                'port': parsed.port or 8080,
                'type': 'http'
            }
        elif 'https' in proxy_handler:
            proxy_url = proxy_handler['https']
            parsed = urllib.parse.urlparse(proxy_url)
            return {
                'host': parsed.hostname,
                'port': parsed.port or 8080,
                'type': 'http'
            }
    except Exception:
        pass
    
    return None


def create_imap_connection(host, port=None, use_ssl=True, proxy_config=None, max_retries=3):
    """创建IMAP连接，支持代理和重试机制"""
    import time
    
    # 如果没有提供proxy_config但需要使用代理，则获取系统代理
    if proxy_config is None:
        # 不使用代理
        pass
    elif proxy_config.get('enabled', False):
        # 使用提供的代理配置
        pass
    else:
        # 代理被禁用
        proxy_config = None
    
    last_error = None
    
    for attempt in range(max_retries):
        try:
            if use_ssl:
                if proxy_config and proxy_config.get('enabled', False):
                    conn = ProxyIMAP4_SSL(
                        host=host, 
                        port=port or 993,
                        proxy_host=proxy_config.get('host'),
                        proxy_port=proxy_config.get('port'),
                        proxy_type=proxy_config.get('type', 'http'),
                        proxy_username=proxy_config.get('username'),
                        proxy_password=proxy_config.get('password')
                    )
                else:
                    conn = imaplib.IMAP4_SSL(host=host, port=port or 993)
            else:
                if proxy_config and proxy_config.get('enabled', False):
                    conn = ProxyIMAP4(
                        host=host, 
                        port=port or 143,
                        proxy_host=proxy_config.get('host'),
                        proxy_port=proxy_config.get('port'),
                        proxy_type=proxy_config.get('type', 'http'),
                        proxy_username=proxy_config.get('username'),
                        proxy_password=proxy_config.get('password')
                    )
                else:
                    conn = imaplib.IMAP4(host=host, port=port or 143)
            
            # 验证连接状态
            try:
                conn.noop()
                return conn
            except Exception as e:
                conn.logout()
                raise e
                
        except (ssl.SSLError, socket.error, Exception) as e:
            last_error = e
            error_msg = str(e).lower()
            
            # 如果是SSL EOF错误，尝试不同的策略
            if 'eof' in error_msg or 'unexpected_eof' in error_msg:
                if attempt < max_retries - 1:
                    # 等待一段时间后重试
                    time.sleep(1 + attempt * 0.5)
                    continue
            
            # 如果是其他SSL错误，在最后一次尝试时抛出更友好的错误信息
            if attempt == max_retries - 1:
                if 'ssl' in error_msg:
                    raise Exception(f"SSL连接失败: {str(e)}。建议尝试：1) 检查网络连接 2) 禁用代理设置 3) 检查防火墙设置")
                else:
                    raise e
    
    # 如果所有重试都失败了
    raise Exception(f"连接失败，已重试{max_retries}次。最后错误: {str(last_error)}")