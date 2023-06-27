# -*- coding:utf-8 -*-
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import subprocess
from datetime import datetime

# IPアドレスを取得する関数
def get_ip_address():
    result = subprocess.run(['ipconfig'], capture_output=True, text=True)
    for line in result.stdout.split('\n'):
        if 'IPv4 アドレス' in line:
            ip_address = line.split(':')[-1].strip()
            return ip_address
    return None

def main():
    # Googleサービスアカウントの認証情報を指定
    path = './accesskey/path/to/accesskey.json'
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)

    # スプレッドシートに接続
    url = 'url of spreadsheet'
    sheet_name = 'name of sheet'
    gs = gspread.authorize(credentials)
    ss = gs.open_by_url(url)
    worksheet = ss.worksheet(sheet_name)

    # IPアドレスを取得
    ip_address = get_ip_address()

    # 更新時刻を取得
    update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # IPアドレスをスプレッドシートに書き込む
    if ip_address:
        worksheet.update('A2', update_time)
        worksheet.update('B2', ip_address)
    else:
        print("IP address was not found.")
        
        
if __name__ == "__main__":
    main()
