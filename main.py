import os
import zipfile
import glob
from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from gspread_dataframe import get_as_dataframe, set_with_dataframe

# 別コードにて直前にZIPファイルをダウンロードする
# ZIPファイルの展開

# ダウンロードフォルダのパスを指定
download_folder = r'ダウンロードフォルダのパス' 
# ダウンロードフォルダ内のZIPファイルを取得
zip_files = glob.glob(os.path.join(download_folder, '*.zip'))

if not zip_files:
    print("ダウンロードフォルダにZIPファイルが見つかりません。")
else:
    # 最新のZIPファイルを取得
    latest_zip_file = max(zip_files, key=os.path.getctime)

    # ZIPファイルを展開
    with zipfile.ZipFile(latest_zip_file) as zip_ref:
        for info in zip_ref.infolist():
            info.filename = info.orig_filename.encode('cp437').decode('cp932')
            if os.sep != "/" and os.sep in info.filename:
                info.filename = info.filename.replace(os.sep, "/")
            zip_ref.extract(info, download_folder)

    # ZIPファイル内のファイルを開く（この例では最初のファイルを開く）
    with zipfile.ZipFile(latest_zip_file, 'r') as zip_ref:
        first_file_in_zip = zip_ref.namelist()[0]
        with zip_ref.open(first_file_in_zip) as file:
            # ここでファイルを使用するための適切な処理を行います
            print(f"{latest_zip_file} を展開し、ファイル名を正しくデコードしました。")

# スプレッドシートへCSVファイルをアップロード

# CSV ファイルのパスを指定
csv_file_path = r"CSVファイルのパス" 

# CSV ファイルを読み込む
data = pd.read_csv(csv_file_path)

# 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定
credentials = Credentials.from_service_account_file(
    r"jsonファイルのパス", scopes=scope)
# OAuth2の資格情報を使用してGoogle APIにログイン
gc = gspread.authorize(credentials)

# スプレッドシートIDを変数に格納
SPREADSHEET_KEY = 'スプレッドシートID'

# スプレッドシート（ブック）を開く
workbook = gc.open_by_key(SPREADSHEET_KEY)

# シートの一覧を取得する。（リスト形式）
worksheets = workbook.worksheets()
print(worksheets)

# シートを開く
worksheet = workbook.worksheet('シート名')

# DataFrameをGoogle Sheetsにアップデート
worksheet.clear()
set_with_dataframe(worksheet, data)
