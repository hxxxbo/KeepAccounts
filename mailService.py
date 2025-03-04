# huangxxxbo@163.com
# PMp6VtP7PrdNbGqZ

import imaplib
import email
from email.header import decode_header

# 邮箱配置
IMAP_SERVER = 'imap.163.com'
USERNAME = 'huangxxxbo@163.com'
PASSWORD = 'PMp6VtP7PrdNbGqZ'  # 不是邮箱密码！

# 连接服务器
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
print("已连接服务器")
mail.login(USERNAME, PASSWORD)
print("已登录")



# 上传客户端身份信息
# 【参考：[【2020可用】Python使用 imaplib imapclient连接网易邮箱提示](https://blog.csdn.net/jony_online/article/details/108638571 )】
imaplib.Commands['ID'] = ('AUTH')
args = ("name", "XXXX", "contact", "XXXX@163.com", "version", "1.0.0", "vendor", "myclient")
typ, dat = mail._simple_command('ID', '("' + '" "'.join(args) + '")')
print(mail._untagged_response(typ, dat, 'ID'))


# 关键修复：必须选择邮箱（如 INBOX）
status, msgs = mail.select('inbox')  # <- 添加这一行！
if status != 'OK':
    print("无法选择收件箱")
    mail.logout()
    exit()

# 现在可以执行 SEARCH
status, messages = mail.search(None, 'UNSEEN')
if status != 'OK' or not messages[0]:
    print("没有未读邮件")
    mail.close()
    mail.logout()
    exit()

# 获取最新一封邮件
latest_email_id = messages[0].split()[-1]
status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
raw_email = msg_data[0][1]

# 解析邮件内容
msg = email.message_from_bytes(raw_email)
subject, encoding = decode_header(msg['Subject'])[0]
if encoding:
    subject = subject.decode(encoding)

print(f"主题: {subject}")
print(f"发件人: {msg.get('From')}")

# 提取正文内容
if msg.is_multipart():
    for part in msg.walk():
        content_type = part.get_content_type()
        if content_type == 'text/plain':
            body = part.get_payload(decode=True).decode()
            print(f"正文:\n{body}")
            break
else:
    body = msg.get_payload(decode=True).decode()
    print(f"正文:\n{body}")

mail.close()
mail.logout()
