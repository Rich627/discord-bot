import discord
import asyncio
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os

with open('config.json', 'r') as config_file:
    config = json.load(config_file)

TOKEN = config["TOKEN"]
CHANNEL_ID = config["CHANNEL_ID"]
SHEET_ID = ''
SHEET_RANGE = 'Sheet1!A:D'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def google_service(api, version):
    creds = None
    with open('google-cred.json', 'r') as creds_file:
        creds_json = json.load(creds_file)
        creds = service_account.Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    service = build(api, version, credentials=creds)
    return service

sheet_service = google_service('sheets', 'v4')

intents = discord.Intents.default()
intents.guilds = True
intents.invites = True
client = discord.Client(intents=intents)

async def create_invite_and_update_sheet():
    await client.wait_until_ready()
    channel = await client.fetch_channel(CHANNEL_ID)

    for i in range(251):
        # 讀取特定單元格的值
        name_email_range = f'Sheet1!A{i+2}:B{i+2}'
        name_email_result = sheet_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=name_email_range).execute()
        name_email_values = name_email_result.get('values', [])
        
        if not name_email_values:
            print("No data found.")
            continue

        # A 列是姓名，B 列是電子郵件
        name, email = name_email_values[0]

        # 創建邀請連結
        invite = await channel.create_invite(max_uses=1, max_age=0)

        # 更新邀請連結到 Google Sheets
        next_row = i + 2 # 假設每次都是寫入新的一行
        new_range = f'Sheet1!E{next_row}'
        body = {'values': [[invite.url]]}
        result = sheet_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID, range=new_range,
            valueInputOption='RAW', body=body).execute()

        # 在 Discord 頻道發送訊息
        message_content = f"{name} {email} {invite.url}"
        await channel.send(message_content)

        # 發送寫入成功訊息
        success_message = "資料寫入成功！"
        await channel.send(success_message)
        
        await asyncio.sleep(1)

    await client.close()

@client.event
async def on_ready():
    print(f'Logged in as {client.user.name}')
    await create_invite_and_update_sheet()

def main():
    client.run(TOKEN)

if __name__ == "__main__":
    main()
