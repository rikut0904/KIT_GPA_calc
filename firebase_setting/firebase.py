import os
from dotenv import load_dotenv
import pyrebase
from function.gui import reload_gui
import firebase_admin
from firebase_admin import firestore, credentials
from datetime import datetime

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
if not firebase_admin._apps:
    firebase = pyrebase.initialize_app(get_firebase_config())
    cred = credentials.Certificate("firebase_setting/firebase_api.json")
    firebase_admin.initialize_app(cred)

# 各サービスのインスタンスを作成
def get_auth():
    return firebase.auth()

def get_database():
    return firestore.client()

# ログイン
def login_function(ls, password, state, win, auth, db):
    print("ログイン")
    try:
        user = auth.sign_in_with_email_and_password(state.user_email, password)
        UserID = user["localId"]
        
        # Firestore にデータを保存・取得
        user_ref = db.collection("users").document(UserID)
        user_data = user_ref.get()
        if user_data.exists:
            UserName = user_data.get("UserName")
            state.update_state(user_name_param=UserName, user_id_param=UserID, idToken_param=user["idToken"])
            ls = firebase_save(ls, state, db)
            print(ls)
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": UserName,
            "action": f"{UserName}がログインしました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        print(f"ログイン成功")
        state.update_state(login=True, p_login=False)
        win = reload_gui(state, win)
        win["-UserName-"].update(state.user_name)
        win["-email-"].update(state.user_email)
    except Exception as e:
        print(f"ログインに失敗しました: {e}")
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": "Guest",
            "action": "ログインに失敗しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        txt = f"ログインに失敗しました。{e}"
        win["-txt-"].update(txt)
        state.update_state(user_name_param="", user_email_param="", user_id_param="", idToken_param="")
    return win, ls

# ユーザー作成
def create_user(ls, password, state, win, auth, db):
    print("ユーザー新規作成")
    try:
        user = auth.create_user_with_email_and_password(state.user_email, password)
        UserID = user["localId"]
        # Firestore にデータを保存
        user_ref = db.collection("users").document(UserID)
        user_data = user_ref.get()
        state.update_state(user_id_param=UserID, idToken_param=user["idToken"])
        if not(user_data.exists):
            user_ref.set({
                "UserName": state.user_name,
                "email": state.user_email,
                "time_stamp": firestore.SERVER_TIMESTAMP,
            })
            ls = firebase_save(ls, state, db)
            print(ls)
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": state.user_name,
            "action": f"{state.user_name}が新規作成しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        print(f"ユーザーが作成されました。")
        state.update_state(login=True, p_login=False, p_signup=False)
        win = reload_gui(state, win)
        win["-UserName-"].update(state.user_name)
        win["-email-"].update(state.user_email)
    except Exception as e:
        print(f"ユーザーの作成に失敗しました: {e}")
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": "Guest",
            "action": "ユーザーの作成に失敗しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        txt = f"ユーザーの作成に失敗しました。{e}"
        win["-txt-"].update(txt)
        state.update_state(user_name_param="", user_email_param="", user_id_param="", idToken_param="")
    return win, ls

# ログアウト
def logout_function(state, win, auth, db):
    print("ログアウト")
    db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
        "user_name": state.user_name,
        "action": f"{state.user_name}がログアウトしました",
        "time_stamp": firestore.SERVER_TIMESTAMP
    })
    auth.current_user = None
    state.update_state(login=False, p_login=False, p_signup=False, user_name_param="", user_email_param="", user_id_param="", idToken_param="")
    win = reload_gui(state, win)
    win["-UserName-"].update(state.user_name)
    win["-email-"].update(state.user_email)
    return win

# ユーザー削除
def delete_user(ls, state, win, auth, db):
    print("ユーザー削除")
    try:
        ls = delete_all_subject(ls, state, db)
        db.collection("users").document(state.user_id).delete()
        auth.delete_user_account(state.idToken)
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": state.user_name,
            "action": f"{state.user_name}を削除しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        print(f"ユーザーが削除されました。")
        state.update_state(login=False, user_name_param="", user_email_param="", user_id_param="", idToken_param="")
        win = reload_gui(state, win)
        win["-UserName-"].update(state.user_name)
        win["-email-"].update(state.user_email)
    except Exception as e:
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": state.user_name,
            "action": f"{state.user_name}がユーザーの削除に失敗しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
        print(f"ユーザーの削除に失敗しました: {e}")
    return win, ls

# firebaseに保存
def firebase_save(ls, state, db):
    print("firebaseに保存")
    user = db.collection("users").document(state.user_id)
    user_sub = user.collection("subject_data")
    user_data = user.get()
    if user_data.exists:
        if len(ls) > 1:
            for data in ls[1:]:
                user.update({
                    "time_stamp": firestore.SERVER_TIMESTAMP
                })
                subject_name, subject_unit, subject_grade, subject_pass_fail = data
                user_sub.document(subject_name).set({
                    "subject_name": subject_name,
                    "subject_unit": subject_unit,
                    "subject_grade": subject_grade,
                    "subject_pass_fail": subject_pass_fail
                })
        db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
            "user_name": state.user_name,
            "action": f"{state.user_name}が成績情報を保存しました",
            "time_stamp": firestore.SERVER_TIMESTAMP
        })
    return firebase_get(state, db)

# firebaseから取得
def firebase_get(state, db):
    print("firebaseから取得")
    ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
    user_ref = db.collection("users").document(state.user_id).collection("subject_data").get()
    if user_ref:
        for doc in user_ref:
            data = doc.to_dict()
            ls.append([data["subject_name"], data["subject_unit"], data["subject_grade"], data["subject_pass_fail"]])
        return ls
    else:
        return [["科目名", "単位数", "評価ポイント", "合否科目"]]

# 科目修正
def update_subject(ls, subject, units_num, HPT, Pass_Fail, state, db):
    print("科目修正")
    if state.isLogin:
        try:
            data = db.collection("users").document(state.user_id).collection("subject_data").document(subject)
            subject_data = data.get()

            if subject_data.exists:
                old_data = subject_data.to_dict()
                old_units_num = old_data.get("subject_unit", "")
                old_HPT = old_data.get("subject_grade", "")
                old_Pass_Fail = old_data.get("subject_pass_fail", "")
                print(old_units_num, old_HPT, old_Pass_Fail)
            if units_num == "":
                units_num = old_units_num
            if HPT == "":
                HPT = old_HPT
            if Pass_Fail == "":
                Pass_Fail = old_Pass_Fail
            data.update({
                "subject_unit": units_num,
                "subject_grade": HPT,
                "subject_pass_fail": Pass_Fail
            })
            db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
                "user_name": state.user_name,
                "action": f"{state.user_name}が成績情報を修正しました",
                "time_stamp": firestore.SERVER_TIMESTAMP
            })
            print(f"科目が修正されました。")
            ls = firebase_get(state, db)
        except Exception as e:
            print(f"科目の修正に失敗しました: {e}")
            db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
                "user_name": state.user_name,
                "action": f"{state.user_name}が成績情報を修正に失敗しました",
                "time_stamp": firestore.SERVER_TIMESTAMP
            })
    else:
        try:
            for data in ls[1:]:
                old_subject, old_units_num, old_HPT, old_Pass_Fail = data
                print(old_subject, old_units_num, old_HPT, old_Pass_Fail)
                if old_subject == subject:
                    if units_num == "":
                        units_num = old_units_num
                    if HPT == "":
                        HPT = old_HPT
                    if Pass_Fail == "":
                        Pass_Fail = old_Pass_Fail
                    data[1], data[2], data[3] = units_num, HPT, Pass_Fail
            print(f"科目が修正されました。")
        except Exception as e:
            print(f"科目の修正に失敗しました: {e}")
    return ls

# 科目一部削除
def delete_subject(ls, subject, state, db):
    print("科目削除")
    if state.isLogin:
        try:
            user = db.collection("users").document(state.user_id)
            user_sub = user.collection("subject_data")
            user_sub.document(subject).delete()
            db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
                "user_name": state.user_name,
                "action": f"{state.user_name}が成績情報を削除しました",
                "time_stamp": firestore.SERVER_TIMESTAMP
            })
            print(f"科目が削除されました。")
            ls = firebase_get(state, db)
        except Exception as e:
            print(f"科目の削除に失敗しました: {e}")
            db.collection("logs").document(f"log_{datetime.now().strftime('%Y-%m-%d_%H:%M:%S')}").set({
                "user_name": state.user_name,
                "action": f"{state.user_name}が成績情報を削除に失敗しました",
                "time_stamp": firestore.SERVER_TIMESTAMP
            })
    else:
        try:
            for data in ls[1:]:
                if data[0] == subject:
                    ls.remove(data)
            print(f"科目が削除されました。")
        except Exception as e:
            print(f"科目の削除に失敗しました: {e}")
    return ls

# 成績情報全削除
def delete_all_subject(ls, state, db):
    print("成績情報全削除")
    try:
        user = db.collection("users").document(state.user_id)
        subjects = user.collection("subject_data").stream()
        for doc in subjects:
            doc.reference.delete()
        
        print(f"成績情報が全て削除されました。")
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
    except Exception as e:
        print(f"成績情報の全削除に失敗しました: {e}")
    return ls