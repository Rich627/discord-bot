import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage



# Google Sheets認證和設置
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/rich/Desktop/discord_bot/gcp-cred.json', scope)
client = gspread.authorize(creds)

# 打開Google Sheets
sheet_id = '11LVgaHo5kts1b_-3tJ7pWWZFbxvgenx7pFxR_yXtXE0'
sheet = client.open_by_key(sheet_id).sheet1

# 讀取所有行
rows = sheet.get_all_records()

# 讀取郵件模板
with open('template1.html', 'r') as file:
    template = file.read()

# 配置SMTP服務器
smtp_server = "smtp.gmail.com"
port = 587  
sender_email = "rich.liu627@gmail.com"
password = "zpog freo wdqz tdtq"

# 建立SMTP連接
server = smtplib.SMTP(smtp_server, port)
server.starttls()  # 啟用安全傳輸
server.login(sender_email, password)

# 為每一行數據發送郵件
for row in rows:
    receiver_email = row['email'] 
    message = MIMEMultipart("related")
    message["Subject"] = "【報名成功】歡迎加入「 6th AWS Educate Taiwan 雲端校園大使證照陪跑計畫」！"
    message["From"] = sender_email
    message["To"] = receiver_email
    
    # 將模板中的佔位符替換為實際數據
    personalized_content = template.format(name=row['name'], link=row['discord link'])
    part = MIMEText(personalized_content, "html")
    message.attach(part)

    # with open("/Users/rich/Desktop/discord_bot/AWS_Educate_logo.png", "rb") as img_file: 
    #     img = MIMEImage(img_file.read())
    #     img.add_header('Content-ID', '<aws_logo>')  
    #     message.attach(img)
    
    # 發送郵件
    server.sendmail(sender_email, receiver_email, message.as_string())

# 斷開SMTP連接
server.quit()
