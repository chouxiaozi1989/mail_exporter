# 163邮箱邮件导出工具

这是一个用于导出163邮箱邮件的Python工具，支持命令行和图形界面两种使用方式。可以按时间段收取指定文件夹中的邮件，并将邮件的收件时间、发件人、主题、内容汇总为CSV文件。

## 功能特点

- 🖥️ **双模式支持**：提供命令行和图形界面两种使用方式
- 📅 **时间段筛选**：支持按开始和结束日期筛选邮件
- 📁 **文件夹选择**：支持选择不同的邮箱文件夹（收件箱、发件箱等）
- 🔄 **实时进度显示**：显示邮件处理进度和状态信息
- 📧 **智能邮件解析**：自动解析邮件日期、发件人、主题和正文内容
- 🌐 **多编码支持**：处理各种编码格式的邮件内容（UTF-8、GBK等）
- 📊 **CSV导出**：导出结果为CSV格式，方便数据分析和处理
- 🔧 **批量处理**：支持大量邮件的分批处理，避免内存问题
- 🛡️ **错误处理**：完善的错误处理和重试机制
- 📝 **详细日志**：提供详细的处理日志和错误信息

## 安装依赖

本工具使用Python 3开发，主要依赖标准库。如果使用图形界面，需要安装tkinter（通常Python自带）。

### 系统要求
- Python 3.6 或更高版本
- tkinter（图形界面需要，通常随Python安装）
- 网络连接（用于IMAP连接）

### 检查依赖
```bash
# 检查Python版本
python --version

# 检查tkinter是否可用
python -c "import tkinter; print('tkinter可用')"
```

## 使用方法

### 方式一：图形界面（推荐）

直接运行GUI程序：
```bash
python mail_exporter_gui.py
```

图形界面功能：
- 📧 **邮箱配置**：输入邮箱地址和授权码
- 📁 **文件夹选择**：从下拉列表中选择要导出的邮箱文件夹
- 📅 **日期设置**：通过日期选择器设置开始和结束日期
- 📂 **文件选择**：选择CSV输出文件的保存位置
- ▶️ **一键导出**：点击开始按钮即可开始导出
- 📊 **实时进度**：显示处理进度条和详细状态信息
- 📝 **操作日志**：查看详细的处理日志和错误信息

### 方式二：命令行参数

```bash
python mail_exporter.py [-h] [-u USERNAME] [-p PASSWORD] [-s START] [-e END] [-o OUTPUT] [-f FOLDER]
```

参数说明：
- `-h, --help`: 显示帮助信息
- `-u, --username`: 邮箱用户名
- `-p, --password`: 邮箱密码或授权码
- `-s, --start`: 开始日期 (YYYY-MM-DD格式)
- `-e, --end`: 结束日期 (YYYY-MM-DD格式)
- `-o, --output`: 输出CSV文件路径，默认为emails.csv
- `-f, --folder`: 邮箱文件夹，默认为INBOX

### 示例

#### 图形界面使用示例
1. 运行 `python mail_exporter_gui.py`
2. 输入邮箱地址：`your_email@163.com`
3. 输入授权码：`your_authorization_code`
4. 点击「刷新文件夹」获取可用文件夹列表
5. 选择要导出的文件夹（如：INBOX）
6. 设置开始和结束日期
7. 选择输出文件路径
8. 点击「开始导出」

#### 命令行使用示例
```bash
# 导出当前月份的所有邮件
python mail_exporter.py -u your_email@163.com -p your_password

# 导出指定时间段的邮件
python mail_exporter.py -u your_email@163.com -p your_password -s 2023-01-01 -e 2023-12-31

# 导出到指定文件
python mail_exporter.py -u your_email@163.com -p your_password -o my_emails.csv

# 导出发件箱邮件
python mail_exporter.py -u your_email@163.com -p your_password -f "Sent Messages"
```

## 注意事项

### 邮箱设置要求
1. **开启IMAP服务**：需要在163邮箱设置中开启IMAP服务
2. **使用授权码**：建议使用应用专用授权码而非邮箱主密码
3. **网络连接**：确保网络连接稳定，避免导出过程中断

### 使用建议
1. **大量邮件**：对于大量邮件的导出，可能需要较长时间，请耐心等待
2. **分批导出**：如果邮件数量很大（>1000封），建议分时间段导出
3. **文件路径**：确保输出路径有足够的磁盘空间
4. **编码问题**：如果CSV文件中文显示异常，请使用支持UTF-8的编辑器打开

### 性能优化
- 工具采用分批处理机制，每批处理50封邮件
- 自动重试机制处理网络不稳定情况
- 连接保活机制避免长时间导出时连接断开

## 授权码获取方法

### 163邮箱授权码设置步骤
1. 登录163邮箱网页版
2. 点击右上角「设置」
3. 选择「POP3/SMTP/IMAP」选项卡
4. 开启「IMAP/SMTP服务」
5. 点击「授权密码管理」
6. 按提示获取授权码（通常需要手机验证）
7. 将获取的授权码作为密码使用

### 其他邮箱服务商
本工具理论上支持其他支持IMAP的邮箱服务商，但可能需要修改IMAP服务器地址。

## 故障排除

### 常见问题

#### 1. 认证失败
**错误信息**：`IMAP错误: [AUTHENTICATIONFAILED]`
**解决方法**：
- 检查邮箱地址和授权码是否正确
- 确认已开启IMAP服务
- 尝试重新获取授权码

#### 2. 连接超时
**错误信息**：`连接超时` 或 `网络错误`
**解决方法**：
- 检查网络连接
- 尝试使用VPN或更换网络环境
- 稍后重试

#### 3. 找不到邮件
**错误信息**：`未找到符合条件的邮件`
**解决方法**：
- 检查日期范围设置
- 确认选择的文件夹中有邮件
- 尝试扩大日期范围

#### 4. CSV文件乱码
**问题描述**：导出的CSV文件中文显示为乱码
**解决方法**：
- 使用支持UTF-8编码的编辑器（如VS Code、Notepad++）
- 在Excel中导入时选择UTF-8编码

### 调试模式
如果遇到问题，可以查看详细的错误日志：
- 图形界面：查看底部的日志区域
- 命令行：查看终端输出的详细信息

## 📦 打包和分发

### 创建可执行文件

项目提供了完整的打包配置，可以将Python程序打包成Windows可执行文件和安装程序。

#### 快速打包

```bash
# 使用自动打包脚本
build.bat

# 或使用Python配置脚本
python build_config.py
```

#### 手动打包

```bash
# 安装PyInstaller
pip install pyinstaller

# 打包GUI程序
pyinstaller --onefile --windowed --name="mail_exporter_gui" mail_exporter_gui.py

# 打包命令行程序
pyinstaller --onefile --name="mail_exporter" mail_exporter.py
```

### 创建安装程序

使用Inno Setup创建Windows安装程序：

1. 下载并安装 [Inno Setup](https://jrsoftware.org/isinfo.php)
2. 运行编译命令：
   ```bash
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
   ```

详细的打包说明请参考 [BUILD.md](BUILD.md) 文档。

## 更新日志

### v1.2.0
- ✨ 新增图形界面支持
- 🔧 修复CSV文件只写入最后一批数据的问题
- 📊 改进进度显示和错误处理
- 🛡️ 增强连接稳定性和重试机制

### v1.1.0
- 🔧 修复6月份邮件获取问题
- 📅 改进日期格式处理
- 📝 增强错误日志和调试信息

### v1.0.0
- 🎉 初始版本发布
- 📧 基本邮件导出功能
- 📊 CSV格式输出
- 📦 添加完整的打包配置和安装程序支持

## 贡献

欢迎提交Issue和Pull Request来改进这个工具！

## 许可证

MIT License