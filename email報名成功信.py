import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# https://www.gdoctohtml.com/
# 設定 Gmail 登入資訊
gmail_user = 'awseducate.cloudambassador@gmail.com'
gmail_password = ""


# 建立 SMTP 連接
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(gmail_user, gmail_password)

# 讀取 Excel 文件
# data = pd.read_excel('/Users/rich/Desktop/discrod_bot/template.xlsx') 
data = pd.read_excel('/Users/rich/Desktop/discrod_bot/陪跑計劃錄取表單_第十三梯次.xlsx') 
# Read the HTML template
with open('/Users/rich/Desktop/discrod_bot/template2.html', 'r') as file:  # Replace with your actual file path
    template = file.read()
# print(data['姓名'])
# Loop through each row in the data to send an email
for index, row in data.iterrows():
    receiver_email = row['Email']  # Make sure this is your column name in Excel
    discord_link = row['Discord Link']  # Make sure this is your column name in Excel
    message = MIMEMultipart("related")
    message["Subject"] = "【報名成功】歡迎加入「 6th AWS Educate Taiwan 雲端校園大使證照陪跑計畫」！"
    message["From"] = gmail_user
    message["To"] = receiver_email
    print(f"Sending to {row['姓名']}")  # Confirm the column name is '姓名'
    
    # Replace the placeholder with actual data
    personalized_content = template.format(name=row['姓名'], discord_link=row['Discord Link'])  # Make sure placeholders match DataFrame column names
    part = MIMEText(personalized_content, "html")
    message.attach(part)
    
    # Send the email
    server.sendmail(gmail_user, receiver_email, message.as_string())

# Disconnect the SMTP connection
server.quit()

