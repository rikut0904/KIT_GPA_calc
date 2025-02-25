import os
from dotenv import load_dotenv
import pyrebase
from function.gui import reload_gui
import firebase_admin
from firebase_admin import firestore, credentials

def get_firebase_config():
    # .envファイルを読み込む
    load_dotenv()
    
    # 環境変数から設定を取得
    firebase_config = {
        "apiKey": os.getenv("REACT_APP_FIREBASE_API_KEY"),
        "authDomain": os.getenv("REACT_APP_FIREBASE_AUTH_DOMAIN"),
        "databaseURL": os.getenv("REACT_APP_DATABASE_URL"),
        "projectId": os.getenv("REACT_APP_DATABASE_PROJECT_ID"),
        "storageBucket": os.getenv("REACT_APP_FIREBASE_STORAGE_BUCKET"),
        "messagingSenderId": os.getenv("REACT_APP_FIREBASE_MESSAGING_SENDER_ID"),
        "appId": os.getenv("REACT_APP_FIREBASE_APP_ID"),
        "measurementId": os.getenv("REACT_APP_FIREBASE_MEASUREMENTID")
    }
    print(firebase_config)
    
    return firebase_config

# Firebaseインスタンスを初期化
firebase = pyrebase.initialize_app(get_firebase_config())
cred = credentials.Certificate("firebase_setting/firebase_api.json")
firebase_admin.initialize_app(cred)

# 各サービスのインスタンスを作成
def get_auth():
    return firebase.auth()

def get_database():
    return firestore.client()

# ログイン
def login_function(password, state, win, auth, db):
    print("ログイン")
    try:
        user = auth.sign_in_with_email_and_password(state.user_email, password)
        UserID = user["localId"]
        
        # Firestore にデータを保存・取得
        user_ref = db.collection("users").document(UserID)
        user_data = user_ref.get()
        if user_data.exists:
            user_ref.update({
                "time_stamp": firestore.SERVER_TIMESTAMP,
            })
            UserName = user_data.get("UserName")
        print(f"ログイン成功")
        state.update_state(login=True, p_login=False, user_name_param=UserName, user_id_param=UserID, idToken_param=user["idToken"])
        win = reload_gui(state, win)
    except Exception as e:
        print(f"ログインに失敗しました: {e}")
        state.update_state(user_email_param="ログインしてください")
    return win

# ユーザー作成
def create_user(password, state, win, auth, db):
    print("ユーザー新規作成")
    try:
        user = auth.create_user_with_email_and_password(state.user_email, password)
        UserID = user["localId"]
        # Firestore にデータを保存
        user_ref = db.collection("users").document(UserID)
        user_data = user_ref.get()
        if not(user_data.exists):
            user_ref.set({
                "UserName": state.user_name,
                "email": state.user_email,
                "time_stamp": firestore.SERVER_TIMESTAMP
            })
        print(f"ユーザーが作成されました。")
        state.update_state(login=True, p_login=False, p_signup=False, user_id_param=UserID, idToken_param=user["idToken"])
        win = reload_gui(state, win)
    except Exception as e:
        print(f"ユーザーの作成に失敗しました: {e}")
        state.update_state(user_name_param="ログインしてください", user_email_param="ログインしてください")
    return win

# ログアウト
def logout_function(state, win, auth):
    print("ログアウト")
    auth.current_user = None
    state.update_state(login=False, p_login=False, p_signup=False, user_name_param="ログインしてください", user_email_param="ログインしてください", user_id_param="ログインしてください", idToken_param="ログインしてください")
    win = reload_gui(state, win)
    return win

# ユーザー削除
def delete_user(state, win, auth, db):
    print("ユーザー削除")
    try:
        auth.delete_user_account(state.idToken)
        db.collection("users").document(state.user_id).delete()
        print(f"ユーザーが削除されました。")
        state.update_state(login=False, user_name_param="ログインしてください", user_email_param="ログインしてください", user_id_param="ログインしてください", idToken_param="ログインしてください")
        win = reload_gui(state, win)
    except Exception as e:
        print(f"ユーザーの削除に失敗しました: {e}")
    return win