import pickle
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# フル権限のスコープを指定
SCOPES = ['https://www.googleapis.com/auth/drive']

def main():
    creds = None
    # 既存の token.pickle があれば削除してから実行することをお勧めします
    if os.path.exists('token.pickle'):
        os.remove('token.pickle')

    # 認証フローの開始
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)

    # 新しいスコープ情報が含まれた token.pickle を保存
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)
    print("新しい token.pickle を作成しました。")

if __name__ == '__main__':
    main()