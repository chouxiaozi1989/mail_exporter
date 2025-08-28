import imaplib
import email
import email.header
import csv
import re
from datetime import datetime, timedelta

def get_mail_folders(username, password):
    """获取邮箱中所有文件夹列表
    
    Args:
        username: 163邮箱用户名
        password: 163邮箱密码或授权码
    
    Returns:
        文件夹列表，格式为 [(folder_name, folder_display_name), ...]
    """
    folders = []
    mail = None
    try:
        # 连接163邮箱IMAP服务器
        mail = imaplib.IMAP4_SSL(host='imap.163.com', port=993)
        mail.login(username, password)
        
        # 获取所有文件夹
        status, folder_list = mail.list()
        
        if status == 'OK':
            for folder_info in folder_list:
                # 解析文件夹信息
                # 格式通常为: (\HasNoChildren) "." "INBOX"
                folder_str = folder_info.decode('utf-8')
                
                # 提取文件夹名称（最后一个引号内的内容）
                import re
                match = re.search(r'"([^"]+)"\s*$', folder_str)
                if match:
                    folder_name = match.group(1)
                    
                    # 创建显示名称（中文化常见文件夹名）
                    display_name = folder_name
                    folder_mapping = {
                        'INBOX': '收件箱',
                        'Sent Messages': '已发送',
                        'Sent': '已发送', 
                        'Drafts': '草稿箱',
                        'Deleted Messages': '已删除',
                        'Trash': '垃圾箱',
                        'Junk': '垃圾邮件',
                        'Spam': '垃圾邮件'
                    }
                    
                    if folder_name in folder_mapping:
                        display_name = f"{folder_mapping[folder_name]} ({folder_name})"
                    
                    folders.append((folder_name, display_name))
        
        # 确保INBOX在第一位
        folders.sort(key=lambda x: (x[0] != 'INBOX', x[1]))
        
    except Exception as e:
        raise Exception(f"获取文件夹列表失败: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except:
                pass
    
    return folders

def fetch_emails(username, password, start_date, end_date, output_csv, folder="INBOX", progress_callback=None):
    """从163邮箱获取指定时间段的邮件并导出为CSV
    
    Args:
        username: 163邮箱用户名
        password: 163邮箱密码或授权码
        start_date: 开始日期，datetime对象
        end_date: 结束日期，datetime对象
        output_csv: 输出CSV文件路径
        folder: 邮箱文件夹，默认为INBOX
        progress_callback: 进度回调函数，接收(current, total, message)参数
    
    Returns:
        导出的邮件数量
    """
    email_count = 0
    try:
        # 连接163邮箱IMAP服务器
        mail = imaplib.IMAP4_SSL(host='imap.163.com', port=993)
        mail.login(username, password)
        
        # 客户端标识（可选）
        imaplib.Commands["ID"] = ('AUTH',)
        args = ("name", username, "contact", username, "version", "1.0.0", "vendor", "myclient")
        mail._simple_command("ID", str(args).replace(",", "").replace("\'", "\""))
        
        # 选择收件箱
        mail.select(folder)
        
        # 格式化日期查询条件 - 强制使用英文locale避免本地化问题
        import locale
        old_locale = locale.getlocale(locale.LC_TIME)
        try:
            # 强制使用C locale确保英文月份缩写
            locale.setlocale(locale.LC_TIME, 'C')
            date_format = '%d-%b-%Y'
            start_date_str = start_date.strftime(date_format)
            end_date_str = (end_date + timedelta(days=1)).strftime(date_format)
            date_query = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
            
            if progress_callback:
                progress_callback(0, 0, f"查询条件: {date_query}")
            else:
                print(f"查询条件: {date_query}")
        finally:
            # 恢复原始locale
            try:
                if old_locale[0]:
                    locale.setlocale(locale.LC_TIME, old_locale)
                else:
                    locale.setlocale(locale.LC_TIME, '')
            except:
                pass
        
        # 搜索邮件 - 使用分批获取解决数量限制
        if progress_callback:
            progress_callback(0, 0, f"开始搜索邮件，查询条件: {date_query}")
        else:
            print(f"开始搜索邮件，查询条件: {date_query}")
        
        status, messages = mail.search(None, date_query)
        
        # 详细的调试信息
        if progress_callback:
            progress_callback(0, 0, f"IMAP搜索状态: {status}, 响应类型: {type(messages)}, 响应长度: {len(messages) if messages else 0}")
        else:
            print(f"IMAP搜索状态: {status}, 响应类型: {type(messages)}, 响应长度: {len(messages) if messages else 0}")
        
        if status != 'OK':
            error_detail = f"搜索邮件失败 - 状态: {status}, 响应: {messages}"
            if progress_callback:
                progress_callback(0, 0, error_detail)
            else:
                print(error_detail)
            raise Exception(error_detail)
        
        # 获取所有邮件ID
        if messages and messages[0]:
            all_message_ids = messages[0].split()
            if progress_callback:
                progress_callback(0, 0, f"原始响应: {messages[0][:200]}{'...' if len(messages[0]) > 200 else ''}")
            else:
                print(f"原始响应: {messages[0][:200]}{'...' if len(messages[0]) > 200 else ''}")
        else:
            all_message_ids = []
            if progress_callback:
                progress_callback(0, 0, "警告: IMAP搜索返回空响应")
            else:
                print("警告: IMAP搜索返回空响应")
        
        total_emails = len(all_message_ids)
        
        if progress_callback:
            progress_callback(0, total_emails, f"找到 {total_emails} 封邮件，开始处理...")
        else:
            print(f"找到 {total_emails} 封邮件，开始处理...")
        
        # 如果没有邮件，直接返回
        if total_emails == 0:
            if progress_callback:
                progress_callback(0, 0, "未找到符合条件的邮件")
                progress_callback(0, 0, "导出完成 - 没有符合条件的邮件")
            else:
                print("未找到符合条件的邮件")
            return 0

        # 准备CSV文件
        with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['日期', '发件人', '主题', '内容'])

            # 分批处理邮件，避免内存问题和服务器限制
            batch_size = 50  # 每批处理50封邮件
            processed_count = 0
            
            for batch_start in range(0, total_emails, batch_size):
                batch_end = min(batch_start + batch_size, total_emails)
                batch_ids = all_message_ids[batch_start:batch_end]
                
                if progress_callback:
                    progress_callback(processed_count, total_emails, f"处理第 {batch_start//batch_size + 1} 批邮件 ({len(batch_ids)} 封)")
                else:
                    print(f"\n处理第 {batch_start//batch_size + 1} 批邮件 ({len(batch_ids)} 封)...")
                
                # 处理当前批次的邮件
                for i, num in enumerate(batch_ids, 1):
                    global_index = processed_count + i
                    
                    # 添加重试机制
                    max_retries = 3
                    retry_count = 0
                    status = 'NO'  # 初始化status变量
                    
                    while retry_count < max_retries:
                        try:
                            status, data = mail.fetch(num, '(RFC822)')
                            if status == 'OK':
                                break
                            else:
                                retry_count += 1
                                if retry_count < max_retries:
                                    import time
                                    time.sleep(0.5)  # 等待0.5秒后重试
                                    continue
                                else:
                                    if progress_callback:
                                        progress_callback(global_index, total_emails, f"跳过邮件 {num} (获取失败)")
                                    else:
                                        print(f"\n警告: 跳过邮件 {num} (获取失败)")
                                    continue
                        except Exception as e:
                            retry_count += 1
                            if retry_count < max_retries:
                                import time
                                time.sleep(0.5)
                                continue
                            else:
                                if progress_callback:
                                    progress_callback(global_index, total_emails, f"跳过邮件 {num} (错误: {str(e)[:50]})")
                                else:
                                    print(f"\n警告: 跳过邮件 {num} (错误: {str(e)[:50]})")
                                break
                    
                    if status != 'OK':
                        continue

                    msg = email.message_from_bytes(data[0][1])
                    # 解析日期
                    try:
                        date_str = msg['Date']
                        if date_str:
                            # 尝试解析各种可能的日期格式
                            from email.utils import parsedate_to_datetime
                            date = parsedate_to_datetime(date_str).strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            date = "[未知日期]"
                    except Exception as e:
                        print(f"解析日期失败: {msg['Date']} - {e}")
                        date = msg['Date'] if msg['Date'] else "[日期解析错误]"
                    from_ = get_mail_from(msg['From'])
                    subject = decode_subject(msg['Subject'])
                    
                    # 获取邮件内容
                    body = ""
                    if msg.is_multipart():
                        # 如果邮件包含多个部分，尝试找到文本部分
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))
                            
                            # 跳过附件
                            if "attachment" in content_disposition:
                                continue
                            
                            # 尝试获取文本内容
                            if content_type == "text/plain":
                                try:
                                    charset = part.get_content_charset() or 'utf-8'
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        try:
                                            body = payload.decode(charset, errors='replace')
                                        except (LookupError, UnicodeDecodeError):
                                            body = payload.decode('utf-8', errors='replace')
                                        break
                                except Exception as e:
                                    print(f"解析邮件内容失败: {e}")
                                    body = "[邮件内容解析错误]"
                    else:
                        # 如果邮件只有一个部分
                        try:
                            charset = msg.get_content_charset() or 'utf-8'
                            payload = msg.get_payload(decode=True)
                            if payload:
                                try:
                                    body = payload.decode(charset, errors='replace')
                                except (LookupError, UnicodeDecodeError):
                                    body = payload.decode('utf-8', errors='replace')
                            else:
                                body = "[空内容]"
                        except Exception as e:
                            print(f"解析邮件内容失败: {e}")
                            body = "[邮件内容解析错误]"
                    
                    # 写入CSV文件 - 对所有邮件都执行
                    writer.writerow([date, from_, subject, body])
                    email_count += 1
                    processed_count += 1  # 每处理完一封邮件就更新计数
                    
                    # 显示处理进度
                    if total_emails > 0:
                        progress = (processed_count / total_emails) * 100
                        if progress_callback:
                            progress_callback(processed_count, total_emails, f"正在处理: {subject[:30]}{'...' if len(subject) > 30 else ''}")
                        else:
                            # 创建进度条
                            bar_length = 30
                            filled_length = int(bar_length * processed_count // total_emails)
                            bar = '=' * filled_length + '-' * (bar_length - filled_length)
                            print(f"\r进度: [{bar}] {processed_count}/{total_emails} ({progress:.1f}%) - {subject[:25]}{'...' if len(subject) > 25 else ''}", end='', flush=True)
                
                # 批次处理完成后的休息和连接保活
                if batch_end < total_emails:
                    import time
                    time.sleep(0.1)  # 批次间短暂休息
                    
                    # 每处理10批邮件后发送NOOP命令保持连接
                    if (batch_start // batch_size + 1) % 10 == 0:
                        try:
                            mail.noop()  # 保持连接活跃
                            if progress_callback:
                                progress_callback(processed_count, total_emails, "保持连接活跃...")
                        except Exception as e:
                            if progress_callback:
                                progress_callback(processed_count, total_emails, f"连接保活失败: {str(e)[:50]}")
                            else:
                                print(f"\n警告: 连接保活失败: {str(e)[:50]}")
    
    except imaplib.IMAP4.error as e:
        error_msg = f"IMAP错误: {e}"
        if progress_callback:
            progress_callback(0, 0, error_msg)
        else:
            print(error_msg)
        raise
    except Exception as e:
        error_msg = f"发生错误: {e}"
        if progress_callback:
            progress_callback(0, 0, error_msg)
        else:
            print(error_msg)
        raise
    finally:
        try:
            mail.close()
            mail.logout()
        except:
            pass
    
    # 确保进度显示后换行
    if not progress_callback:
        if email_count > 0:
            print()  # 换行
        print(f"成功导出 {email_count} 封邮件到 {output_csv}")
    else:
        # 确保进度回调显示完成状态
        if email_count > 0:
            progress_callback(email_count, email_count, f"导出完成 - 共处理 {email_count} 封邮件")
    return email_count

def decode_subject(subject):
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
            print(f"解码主题失败: {subject} - {e}")
            return subject
    else:
        return "[无主题]"

def get_mail_from(from_):
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
            import re
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
                        print(f"无法从 '{from_str}' 中提取邮箱地址")
                        return from_str
        except Exception as e:
            print(f"解析发件人失败: {from_} - {e}")
            return from_
    else:
        return "[未知发件人]"


if __name__ == "__main__":
    import argparse
    import getpass
    import sys
    
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description="163邮箱邮件导出工具")
    parser.add_argument("-u", "--username", help="邮箱用户名")
    parser.add_argument("-p", "--password", help="邮箱密码或授权码")
    parser.add_argument("-s", "--start", help="开始日期 (YYYY-MM-DD)")
    parser.add_argument("-e", "--end", help="结束日期 (YYYY-MM-DD)")
    parser.add_argument("-o", "--output", default="emails.csv", help="输出CSV文件路径")
    parser.add_argument("-f", "--folder", default="INBOX", help="邮箱文件夹，默认为INBOX")
    
    args = parser.parse_args()
    
    # 如果没有提供用户名，提示用户输入
    if not args.username:
        args.username = input("请输入邮箱用户名: ")
        # args.username = 'sdszcsb@163.com'
    
    # 如果没有提供密码，安全提示输入
    if not args.password:
        args.password = getpass.getpass("请输入邮箱密码或授权码: ")
        # args.password = 'LLfBu9hssaME3uXf'
    
    # 处理日期
    try:
        if args.start:
            start_date = datetime.strptime(args.start, "%Y-%m-%d")
        else:
            # 默认为当前月份第一天
            today = datetime.now()
            start_date = datetime(today.year, today.month, 1)
            
        if args.end:
            end_date = datetime.strptime(args.end, "%Y-%m-%d")
        else:
            # 默认为当前日期
            end_date = datetime.now()
    except ValueError as e:
        print(f"日期格式错误: {e}")
        print("请使用YYYY-MM-DD格式")
        sys.exit(1)
    
    print(f"\n开始导出邮件...")
    print(f"用户: {args.username}")
    print(f"时间范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")
    print(f"输出文件: {args.output}")
    print(f"邮箱文件夹: {args.folder}\n")
    
    # 执行邮件导出
    try:
        email_count = fetch_emails(args.username, args.password, start_date, end_date, args.output, args.folder)
        if email_count > 0:
            print(f"\n导出完成! 共导出 {email_count} 封邮件到 {args.output}")
        else:
            print("\n未找到符合条件的邮件")
    except Exception as e:
        print(f"\n导出过程中发生错误: {e}")
        sys.exit(1)