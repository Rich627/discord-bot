import discord
import asyncio
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os

with open('config.json', 'r') as config_file:
    config = json.load(config_file)

TOKEN = config["TOKEN"]
CHANNEL_ID = config["CHANNEL_ID"]
SHEET_ID = '11LVgaHo5kts1b_-3tJ7pWWZFbxvgenx7pFxR_yXtXE0'
SHEET_RANGE = '工作表1!C:C'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_sheets_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'service_account_key.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('sheets', 'v4', credentials=creds)
    return service.spreadsheets()

intents = discord.Intents.default()
intents.guilds = True
intents.invites = True
client = discord.Client(intents=intents)

async def create_invite_and_update_sheet():
    await client.wait_until_ready()
    channel = await client.fetch_channel(CHANNEL_ID)
    sheet_service = get_sheets_service()

    for i in range(20):
        # 讀取特定單元格的值
        name_email_range = f'工作表1!A{i+2}:B{i+2}'
        name_email_result = sheet_service.values().get(spreadsheetId=SHEET_ID, range=name_email_range).execute()
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
        new_range = f'工作表1!C{next_row}'
        body = {'values': [[invite.url]]}
        result = sheet_service.values().update(
            spreadsheetId=SHEET_ID, range=new_range,
            valueInputOption='RAW', body=body).execute()

        # # 在 Discord 頻道發送訊息
        # message_content = f"{name} {email} {invite.url}"
        # await channel.send(message_content)

        # # 發送寫入成功訊息
        # success_message = "資料寫入成功！"
        # await channel.send(success_message)

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
