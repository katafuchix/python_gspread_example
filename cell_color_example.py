import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# サービスアカウントの資格情報ファイルのパス
SERVICE_ACCOUNT_FILE = 'api-project-920307201456-8dd931511e18.json'

# 認証スコープ
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# 資格情報の設定
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Google Sheets APIクライアントの作成
service = build('sheets', 'v4', credentials=creds)

# gspreadでスプレッドシートに接続
gc = gspread.authorize(creds)

# スプレッドシートを開く
spreadsheet_id = '1k7IhCy8AHQD4wtLH886DQwNTjocA8ApWT9sOJoAgXaU'
worksheet = gc.open_by_key(spreadsheet_id).sheet1

# ここではセルA1に書き込む
worksheet.update_acell('A1', 'セルに色を付ける')


# 色付けのためのフォーマット
body = {
    "requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.0,
                            "blue": 0.0
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }
    ]
}

# Google Sheets APIを使ってリクエストを送信
service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body=body
).execute()

print("色を付けました。")
