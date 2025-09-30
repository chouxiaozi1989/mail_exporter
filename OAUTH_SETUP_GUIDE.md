# Gmail OAuth2设置指南

本指南将帮助您设置Gmail OAuth2认证，以便安全地访问Gmail邮箱进行邮件导出。

## 为什么使用OAuth2？

- 🔐 **更安全**：无需使用应用密码，使用Google官方认证协议
- 🛡️ **权限控制**：只授予必要的邮件访问权限
- 🔄 **可撤销**：可随时在Google账户中撤销应用访问权限
- 📱 **现代标准**：符合现代安全认证标准

## 设置步骤

### 第一步：创建Google Cloud项目

1. 访问 [Google Cloud Console](https://console.cloud.google.com/)
2. 点击项目选择器，创建新项目或选择现有项目
3. 为项目命名（例如："邮件导出工具"）

### 第二步：启用Gmail API

1. 在Google Cloud Console中，转到"API和服务" > "库"
2. 搜索"Gmail API"
3. 点击"Gmail API"并点击"启用"

### 第三步：创建OAuth2凭据

1. 转到"API和服务" > "凭据"
2. 点击"创建凭据" > "OAuth客户端ID"
3. 如果首次创建，需要先配置OAuth同意屏幕：
   - 选择"外部"用户类型
   - 填写应用名称（例如："邮件导出工具"）
   - 添加您的邮箱地址作为开发者联系信息
   - 在"范围"部分，添加 `https://mail.google.com/` 范围
4. 选择应用类型为"桌面应用程序"
5. 为OAuth客户端命名（例如："邮件导出客户端"）
6. 点击"创建"

### 第四步：获取客户端凭据

创建完成后，您将获得：
- **客户端ID**：类似 `123456789-abcdefg.apps.googleusercontent.com`
- **客户端密钥**：类似 `GOCSPX-abcdefghijklmnop`

⚠️ **重要**：请妥善保管这些凭据，不要分享给他人。

## 在程序中使用OAuth2

### 第一步：启动程序

运行邮件导出工具的图形界面：
```bash
python mail_exporter_gui.py
```

### 第二步：配置Gmail邮箱

1. 在"邮箱地址"字段输入您的Gmail地址（例如：`your_email@gmail.com`）
2. 程序会自动识别这是Gmail邮箱并显示OAuth2认证选项

### 第三步：输入OAuth2凭据

1. 在"Client ID"字段输入您从Google Cloud Console获得的客户端ID
2. 在"Client Secret"字段输入您的客户端密钥
3. 点击"开始OAuth授权"按钮

### 第四步：完成授权

1. 程序会自动打开浏览器并跳转到Google授权页面
2. 登录您的Google账户
3. 确认授权邮件导出工具访问您的Gmail
4. 授权成功后，浏览器会显示成功页面
5. 返回程序，您会看到"OAuth2已授权"的状态提示

### 第五步：开始导出邮件

1. 授权成功后，Client ID和Secret字段会自动隐藏
2. 选择要导出的邮箱文件夹
3. 设置日期范围
4. 选择输出文件路径
5. 点击"开始导出"即可开始导出邮件

## 重新授权

如果需要切换Google账户或重新授权：

1. 点击"重新授权"按钮
2. Client ID和Secret字段会重新显示
3. 输入新的凭据或相同的凭据
4. 重新完成授权流程

## 故障排除

### 常见问题及解决方案

#### 1. 授权失败：范围错误
**错误信息**：`invalid_scope` 或权限不足
**解决方法**：
- 确保在Google Cloud Console的OAuth同意屏幕中添加了 `https://mail.google.com/` 范围
- 重新创建OAuth客户端凭据

#### 2. 客户端ID无效
**错误信息**：`invalid_client`
**解决方法**：
- 检查客户端ID是否完整复制
- 确保客户端类型为"桌面应用程序"

#### 3. 浏览器无法打开
**解决方法**：
- 手动复制程序显示的授权URL到浏览器
- 完成授权后，将重定向URL中的授权码复制回程序

#### 4. 网络连接问题
**解决方法**：
- 检查网络连接
- 如果在中国大陆，可能需要使用VPN访问Google服务

### 安全注意事项

1. **保护凭据**：不要将Client ID和Secret分享给他人
2. **定期检查**：定期在Google账户中检查已授权的应用
3. **撤销访问**：如不再使用，可在Google账户设置中撤销应用访问权限

## 技术细节

- **OAuth2范围**：`https://mail.google.com/` - 提供完整的Gmail IMAP/SMTP访问权限
- **令牌存储**：授权令牌会安全存储在本地 `gmail_token.json` 文件中
- **自动刷新**：程序会自动处理访问令牌的刷新，无需手动重新授权
- **安全协议**：使用PKCE (Proof Key for Code Exchange) 增强安全性