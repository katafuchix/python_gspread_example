import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from gspread.utils import a1_range_to_grid_range


# シート全体の色をリセットする関数
def clear_sheet(worksheet):
    worksheet.clear()
    # シート全体の色をリセットするリクエスト
    clear_color_request = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": worksheet.id
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            }
        ]
    }

    # Google Sheets APIを使ってリクエストを送信
    service.spreadsheets().batchUpdate(
        spreadsheetId=worksheet.spreadsheet.id,
        body=clear_color_request
    ).execute()


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

clear_sheet(worksheet)

# 複数のセルにテキストを設定
#worksheet.update('A1:B2', [['セル1', 'セル2'], ['セル3', 'セル4']])
worksheet.update(range_name='A1:B2', values=[['セル1', 'セル2'], ['セル3', 'セル4']])

"""
# 色付けのためのフォーマット (A1〜B2の範囲に適用)
body = {
    "requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,  # A1の行（0から始まる）
                    "endRowIndex": 2,    # B2の行（2行目で終了）
                    "startColumnIndex": 0, # A列
                    "endColumnIndex": 2    # B列
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.6,
                            "green": 0.8,
                            "blue": 1.0
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }
    ]
}
"""
"""
# 各セルに異なる色を適用するためのリクエスト index
body = {
    "requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,  # A1の行
                    "endRowIndex": 1,
                    "startColumnIndex": 0, # A列
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.8,
                            "blue": 0.8
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,  # A2の行
                    "endRowIndex": 1,
                    "startColumnIndex": 1, # B列
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.8,
                            "green": 1.0,
                            "blue": 0.8
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,  # B1の行
                    "endRowIndex": 2,
                    "startColumnIndex": 0, # A列
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.8,
                            "green": 0.8,
                            "blue": 1.0
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,  # B2の行
                    "endRowIndex": 2,
                    "startColumnIndex": 1, # B列
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 0.8
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }
    ]
}
"""

# セル範囲をA1形式からインデックスに変換
ranges = ['A1', 'B1', 'A2', 'B2']
colors = [
    {"red": 1.0, "green": 0.8, "blue": 0.8},  # A1
    {"red": 0.8, "green": 1.0, "blue": 0.8},  # B1
    {"red": 0.8, "green": 0.8, "blue": 1.0},  # A2
    {"red": 1.0, "green": 1.0, "blue": 0.8},  # B2
]

# フォーマットリクエストを作成
requests = []
for i, cell in enumerate(ranges):
    grid_range = a1_range_to_grid_range(cell, worksheet.id)  # A1形式をインデックスに変換
    requests.append({
        "repeatCell": {
            "range": grid_range,
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": colors[i]
                }
            },
            "fields": "userEnteredFormat.backgroundColor"
        }
    })

# Google Sheets APIを使ってリクエストを送信
body = {
    "requests": requests
}

# Google Sheets APIを使ってリクエストを送信
service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body=body
).execute()

print("複数のセルに色を付けました。")
