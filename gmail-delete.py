# Gmail を日付指定で削除（ゴミ箱へ移動）する
# 本プログラムで削除実行後はゴミ箱を開いて「ゴミ箱を空にする」を実行すること。
# コマンドラインオプション
# 
# 使い方: python gmail-delete.py [display|delete|count|dispquery] [query='検索条件']
#         display   = 対象のメールを表示
#         delete    = メールを削除
#         count     = 対象のメールの数を表示
#         dispquery = クエリ文字列を表示して終了
#
# 導入:
#         pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# 
import os
import sys
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Gmail API のスコープ
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def process_emails(service, query, mode='display'):
    print(f"検索条件: {query}")

    # count モード（高速件数取得）
    if mode == 'count':
        try:
            next_page_token = None
            total_count = 0
            page_count = 0
            while True:
                page_count += 1
                print(f"\rPage: {page_count}", end='', flush=True)
                response = service.users().messages().list(
                    userId='me',
                    q=query,
                    pageToken=next_page_token,
                    maxResults=500
                ).execute()

                messages = response.get('messages', [])
                total_count += len(messages)

                next_page_token = response.get('nextPageToken')
                if not next_page_token:
                    break

            print(f"\n{total_count} 件のメールを検出しました。")
        except KeyboardInterrupt:
            print("\nカウント処理を中断しました。")
            sys.exit(0)
        return

    # display / delete モード
    try:
        next_page_token = None
        total_count = 0
        page_count = 0

        while True:
            page_count += 1
            print(f"Page: {page_count}")

            response = service.users().messages().list(
                userId='me',
                q=query,
                pageToken=next_page_token,
                maxResults=100
            ).execute()

            messages = response.get('messages', [])
            if not messages:
                if total_count == 0:
                    print("該当するメールはありません。")
                break

            for msg in messages:
                if mode == 'delete':
#                   service.users().messages().delete(userId='me', id=msg['id']).execute()
                    service.users().messages().trash(userId='me', id=msg['id']).execute()
                elif mode == 'display':
                    msg_detail = service.users().messages().get(userId='me', id=msg['id']).execute()
                    headers = msg_detail.get('payload', {}).get('headers', [])
                    subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '(件名なし)')
                    date = next((h['value'] for h in headers if h['name'] == 'Date'), '(日付なし)')
                    print(f"▶ 件名: {subject}\n   日付: {date}\n")
                total_count += 1
                time.sleep(0.01)

            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break

        print(f"\n{total_count} 件のメールを{'削除' if mode == 'delete' else '表示'}しました。")

    except KeyboardInterrupt:
        print("\n処理を中断しました。途中までの件数:", total_count)
        sys.exit(0)

def main():
    try:
        if len(sys.argv) < 2 or sys.argv[1] not in ['display', 'delete', 'count', 'dispquery']:
            print("使い方: python gmail-delete.py [display|delete|count|dispquery] [query='検索条件']")
            print("display   = 対象のメールを表示")
            print("delete    = メールを削除")
            print("count     = 対象のメールの数を表示")
            print("dispquery = クエリ文字列を表示して終了")
            return

        mode = sys.argv[1]
        query = 'in:anywhere before:2025/01/01'  # デフォルト検索条件

        for arg in sys.argv[2:]:
            if arg.startswith("query="):
                query = arg[len("query="):].strip("'\"")

        if mode == 'dispquery':
            print(f"現在の検索クエリ: {query}")
            return

        creds = authenticate()
        service = build('gmail', 'v1', credentials=creds)
        process_emails(service, query, mode)

    except KeyboardInterrupt:
        print("\n起動中に中断されました。")
        sys.exit(0)

if __name__ == '__main__':
    main()
