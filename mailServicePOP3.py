
# huangxxxbo@163.com
# PMp6VtP7PrdNbGqZ


import poplib
from email import parser

# 邮箱配置
POP3_SERVER = 'pop.163.com'
USERNAME = 'huangxxxbo@163.com'
PASSWORD = 'PMp6VtP7PrdNbGqZ'  # 不是邮箱密码！

# 连接服务器
mail = poplib.POP3_SSL(POP3_SERVER)
mail.user(USERNAME)
mail.pass_(PASSWORD)

# 获取邮件数量
num_emails = len(mail.list()[1])

# 获取最新邮件
raw_email = b'\n'.join(mail.retr(num_emails)[1])
msg = parser.BytesParser().parsebytes(raw_email)

print(f"主题: {msg['Subject']}")
print(f"发件人: {msg['From']}")

# 提取正文
if msg.is_multipart():
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            body = part.get_payload(decode=True).decode()
            print(f"正文:\n{body}")
            break
else:
    body = msg.get_payload(decode=True).decode()
    print(f"正文:\n{body}")

mail.quit()


