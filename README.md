# 163邮箱邮件导出工具

这是一个用于导出163邮箱邮件的Python工具，可以按时间段收取收件箱中的邮件，并将邮件的收件时间、发件人、主题、内容汇总为CSV文件。

## 功能特点

- 支持按时间段筛选邮件
- 自动解析邮件日期、发件人、主题和正文
- 处理各种编码格式的邮件内容
- 支持命令行参数配置
- 导出结果为CSV格式，方便后续处理

## 安装依赖

本工具使用Python 3开发，依赖标准库，无需安装额外依赖。

## 使用方法

### 命令行参数

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

```bash
# 导出当前月份的所有邮件
python mail_exporter.py -u your_email@163.com -p your_password

# 导出指定时间段的邮件
python mail_exporter.py -u your_email@163.com -p your_password -s 2023-01-01 -e 2023-12-31

# 导出到指定文件
python mail_exporter.py -u your_email@163.com -p your_password -o my_emails.csv
```

## 注意事项

1. 需要在163邮箱设置中开启IMAP服务
2. 建议使用应用专用密码而非邮箱主密码
3. 如果遇到认证失败，请检查密码是否正确，以及是否已开启IMAP服务
4. 对于大量邮件的导出，可能需要较长时间

## 授权码获取方法

1. 登录163邮箱网页版
2. 点击「设置」-「POP3/SMTP/IMAP」
3. 开启IMAP服务
4. 点击「授权密码管理」获取授权码

## 许可证

MIT