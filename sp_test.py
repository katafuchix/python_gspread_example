import gspread
from google.oauth2.service_account import Credentials

# 認証情報のセットアップ
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# spreadsheets@api-project-920307201456.iam.gserviceaccount.com
# ダウンロードしたサービスアカウントのJSONファイルへのパスを指定
creds = Credentials.from_service_account_file('api-project-920307201456-68526c95dce1.json', scopes=SCOPE)

# gspreadを使ってGoogleスプレッドシートにアクセス
client = gspread.authorize(creds)


id = "1SeDDV6VK3ZcjdQrgsGxaE6M7V_F1VDuvNN692L49aZg"
# GoogleスプレッドシートのIDを指定して開く
spreadsheet = client.open_by_key(id)

# シートの指定（シート名が 'Sheet1' の場合）
sheet = spreadsheet.worksheet('シート1')

print(sheet)

# スプレッドシートの内容を取得（全てのデータ）
data = sheet.get_all_records()

# データの表示
print(data)
