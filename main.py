import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from socket import *
import ssl
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import os

# 全局变量
current_username = ""
current_password = ""
current_host = ""
attachment_paths = []


def login():
    global current_username, current_password, current_host
    current_username = username_entry.get()
    current_password = password_entry.get()
    current_host = host_entry.get()
    messagebox.showinfo("登录", "登录成功！")
    login_frame.pack_forget()
    main_frame.pack()


def send_email(to_entry, subject_entry, message_entry):
    global current_username, current_password, current_host
    toAddress = to_entry.get()
    subject = subject_entry.get()
    msg_body = message_entry.get("1.0", tk.END)

    if not toAddress or not subject or not msg_body:
        messagebox.showerror("错误", "请填写所有必填字段！")
        return

    try:
        # 创建MIMEMultipart对象
        msg = MIMEMultipart()
        msg['From'] = current_username
        msg['To'] = toAddress
        msg['Subject'] = subject

        # 添加邮件正文
        msg.attach(MIMEText(msg_body, 'plain'))

        # 处理附件
        for path in attachment_paths:
            part = MIMEBase('application', 'octet-stream')
            with open(path, 'rb') as attachment_file:
                part.set_payload(attachment_file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(path)}')
            msg.attach(part)

        # 连接到SMTP服务器
        with socket(AF_INET, SOCK_STREAM) as clientSocket:
            clientSocket.connect((current_host, 587))
            recv = clientSocket.recv(1024).decode()
            if not recv.startswith('220'):
                raise Exception('未收到服务器的 220 回复。')

            clientSocket.send('HELO Alice\r\n'.encode())
            recv1 = clientSocket.recv(1024).decode()
            if not recv1.startswith('250'):
                raise Exception('未收到服务器的 250 回复。')

            clientSocket.send('STARTTLS\r\n'.encode())
            recvTLS = clientSocket.recv(1024).decode()
            if not recvTLS.startswith('220'):
                raise Exception('未收到服务器的 220 回复。')

            context = ssl.create_default_context()
            with context.wrap_socket(clientSocket, server_hostname=current_host) as secureSocket:
                secureSocket.sendall('AUTH LOGIN\r\n'.encode())
                recv2 = secureSocket.recv(1024).decode()
                if not recv2.startswith('334'):
                    raise Exception('未收到服务器的 334 回复。')

                secureSocket.sendall(base64.b64encode(current_username.encode()) + b'\r\n')
                recvName = secureSocket.recv(1024).decode()
                if not recvName.startswith('334'):
                    raise Exception('未收到服务器的 334 回复。')

                secureSocket.sendall(base64.b64encode(current_password.encode()) + b'\r\n')
                recvPass = secureSocket.recv(1024).decode()
                if not recvPass.startswith('235'):
                    raise Exception('未收到服务器的 235 回复。')

                secureSocket.sendall(f'MAIL FROM: <{current_username}>\r\n'.encode())
                recvFrom = secureSocket.recv(1024).decode()
                if not recvFrom.startswith('250'):
                    raise Exception('未收到服务器的 250 回复。')

                secureSocket.sendall(f'RCPT TO: <{toAddress}>\r\n'.encode())
                recvTo = secureSocket.recv(1024).decode()
                if not recvTo.startswith('250'):
                    raise Exception('未收到服务器的 250 回复。')

                secureSocket.send('DATA\r\n'.encode())
                recvData = secureSocket.recv(1024).decode()
                if not recvData.startswith('354'):
                    raise Exception('未收到服务器的 354 回复。')

                # 发送完整邮件
                secureSocket.sendall(msg.as_string().encode())

                secureSocket.sendall("\r\n.\r\n".encode())
                recvEnd = secureSocket.recv(1024).decode()
                if not recvEnd.startswith('250'):
                    raise Exception('未收到服务器的 250 回复。')

        # 存储发送记录
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open('sent_emails.txt', 'a', encoding='utf-8') as f:
            f.write(f'发送于: {current_time}\n')
            f.write(f'发件邮箱: {current_username}\n')
            f.write(f'收件邮箱: {toAddress}\n')
            f.write(f'主题: {subject}\n')
            f.write(f'邮件正文: {msg_body}\n')
            f.write('-' * 50 + '\n')
        messagebox.showinfo("成功", "邮件发送成功!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def attach_file():
    global attachment_paths
    attachment_path = filedialog.askopenfilename()
    if attachment_path:
        attachment_paths.append(attachment_path)
        update_attachment_list()


def update_attachment_list():
    for item in attachment_tree.get_children():
        attachment_tree.delete(item)
    for path in attachment_paths:
        attachment_tree.insert("", "end", values=(os.path.basename(path),))

def view_sent_emails():
    try:
        with open('sent_emails.txt', 'r', encoding='utf-8') as f:
            content = f.read()
        messagebox.showinfo("已发送邮件", content)
    except FileNotFoundError:
        messagebox.showinfo("已发送邮件", "已发送邮件为空！")
    except UnicodeDecodeError as e:
        messagebox.showerror("Error", f"无法解码文件，请检查文件编码是否为 utf-8。错误信息: {str(e)}")


def save_to_drafts(to_entry, subject_entry, message_entry):
    toAddress = to_entry.get()
    subject = subject_entry.get()
    msg_body = message_entry.get("1.0", tk.END)
    with open('drafts.txt', 'a') as f:
        f.write(f'收件邮箱: {toAddress}\n')
        f.write(f'主题: {subject}\n')
        f.write(f'邮件正文: {msg_body}\n')
        if attachment_paths:
            f.write(f'附件: {", ".join(attachment_paths)}\n')
        f.write('-' * 50 + '\n')
    messagebox.showinfo("草稿箱", "已存入草稿箱！")


def view_drafts():
    try:
        with open('drafts.txt', 'r') as f:
            content = f.read()
        messagebox.showinfo("草稿箱", content)
    except FileNotFoundError:
        messagebox.showinfo("草稿箱", "草稿箱为空！")


root = tk.Tk()
root.title("邮件客户端")

style = ttk.Style()
style.configure('TButton', font=('Arial', 12))
style.configure('TLabel', font=('Arial', 12))
style.configure('TEntry', font=('Arial', 12))
style.configure('Treeview', font=('Arial', 12), rowheight=30)

login_frame = ttk.Frame(root, padding="10 10 10 10")
login_frame.pack(padx=20, pady=20)

ttk.Label(login_frame, text="邮箱地址:").grid(row=0, column=0, pady=5, sticky=tk.W)
username_entry = ttk.Entry(login_frame, width=50)
username_entry.grid(row=0, column=1, pady=5)
ttk.Label(login_frame, text="密码:").grid(row=1, column=0, pady=5, sticky=tk.W)
password_entry = ttk.Entry(login_frame, width=50, show='*')
password_entry.grid(row=1, column=1, pady=5)
ttk.Label(login_frame, text="邮件服务器地址:").grid(row=2, column=0, pady=5, sticky=tk.W)
host_entry = ttk.Entry(login_frame, width=50)
host_entry.grid(row=2, column=1, pady=5)
ttk.Button(login_frame, text="登录", command=login).grid(row=3, column=1, pady=10)

main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.pack_forget()

ttk.Label(main_frame, text="接受方邮箱:").grid(row=0, column=0, pady=5, sticky=tk.W)
to_entry = ttk.Entry(main_frame, width=50)
to_entry.grid(row=0, column=1, pady=5)
ttk.Label(main_frame, text="邮件主题:").grid(row=1, column=0, pady=5, sticky=tk.W)
subject_entry = ttk.Entry(main_frame, width=50)
subject_entry.grid(row=1, column=1, pady=5)
ttk.Label(main_frame, text="邮件正文:").grid(row=2, column=0, pady=5, sticky=tk.NW)
message_entry = tk.Text(main_frame, height=10, width=50, font=('Arial', 12))
message_entry.grid(row=2, column=1, pady=5)

button_frame = ttk.Frame(main_frame)
button_frame.grid(row=3, column=0, columnspan=2, pady=10)

attachment_tree = ttk.Treeview(main_frame, columns=("Filename",), show="headings", height=5)
attachment_tree.heading("Filename", text="附件名称")
attachment_tree.grid(row=4, column=1, pady=5)

ttk.Button(button_frame, text="发送邮件", command=lambda: send_email(to_entry, subject_entry, message_entry)).pack(
    side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="查看已发送邮件", command=view_sent_emails).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="保存到草稿箱",
           command=lambda: save_to_drafts(to_entry, subject_entry, message_entry)).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="查看草稿箱", command=view_drafts).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="添加附件", command=attach_file).pack(side=tk.LEFT, padx=10)

root.mainloop()