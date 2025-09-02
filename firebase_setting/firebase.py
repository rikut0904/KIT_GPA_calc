import firebase_admin
from firebase_admin import credentials, auth, firestore
import os
from dotenv import load_dotenv
from function import state
from function.gui import reload_gui
import tkinter as tk
from tkinter import messagebox

# 環境変数を読み込み
load_dotenv()

def load_from_local_csv(user_name):
    """ローカルCSVファイルからデータを読み込み"""
    try:
        import csv
        filename = f"subject_grades_data_{user_name}.csv"
        print(f"ローカルCSVファイルを読み込み中: {filename}")
        
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                ls = list(reader)
            print(f"ローカルCSVから読み込み成功: {len(ls)}行")
            return ls
        else:
            print(f"ローカルCSVファイルが見つかりません: {filename}")
            return [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
    except Exception as e:
        print(f"ローカルCSV読み込みエラー: {e}")
        return [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]

# Firebase設定
def get_auth():
    """Firebase Authentication サービスを取得"""
    try:
        return auth
    except Exception as e:
        print(f"Firebase Auth の初期化に失敗しました: {e}")
        return None

def get_database():
    """Firestore データベースサービスを取得"""
    try:
        return firestore.client()
    except Exception as e:
        print(f"Firestore の初期化に失敗しました: {e}")
        return None

def initialize_firebase():
    """Firebase を初期化"""
    try:
        # 既に初期化されている場合はスキップ
        if firebase_admin._apps:
            return True
            
        # 環境変数から設定を取得
        project_id = os.getenv('REACT_APP_DATABASE_PROJECT_ID')
        
        # 設定が不完全な場合はエラーを返す
        if not project_id:
            print("Firebase設定が不完全です。")
            print("実際のFirebase機能を使用するには、.envファイルに正しい設定を追加してください。")
            print("必要な設定:")
            print("- REACT_APP_DATABASE_PROJECT_ID")
            return False
            
        # サービスアカウントキーファイルを使用
        service_account_path = os.path.join(os.path.dirname(__file__), 'firebase_api.json')
        
        if os.path.exists(service_account_path):
            # サービスアカウントキーファイルを使用
            cred = credentials.Certificate(service_account_path)
            firebase_admin.initialize_app(cred, {
                'projectId': project_id,
            })
            print("Firebase が正常に初期化されました（サービスアカウントキー使用）")
            return True
        else:
            print("サービスアカウントキーファイル（firebase_api.json）が見つかりません。")
            print("Firebase Console > プロジェクト設定 > サービスアカウント から、サービスアカウントキーをダウンロードし、")
            print("firebase_setting/firebase_api.jsonとして保存してください。")
            return False
        
    except Exception as e:
        print(f"Firebase の初期化に失敗しました: {e}")
        return False

def login_function(ls, password, state, win, auth_service, db):
    """Firebase認証を使用したログイン機能"""
    try:
        email = state.current_user_email
        
        # 基本的なバリデーション
        if not email or not password:
            messagebox.showerror("エラー", "メールアドレスとパスワードを入力してください。")
            return win, ls
        
        if "@" not in email:
            messagebox.showerror("エラー", "有効なメールアドレスを入力してください。")
            return win, ls
        
        # Firebase認証が利用可能でない場合
        if not auth_service or not db:
            messagebox.showerror("エラー", "Firebase認証サービスが利用できません。\nFirebase設定を確認してください。")
            return win, ls
        
        # Firebase Authenticationを使用してログイン
        try:
            # Firebase Admin SDKを使用してユーザーを取得
            user = auth_service.get_user_by_email(email)
            
            # ユーザーが存在する場合
            if user:
                # ログイン成功
                user_name = user.display_name or email.split("@")[0]
                
                state.update_state(
                    login=True,
                    user_name_param=user_name,
                    user_email_param=email,
                    user_id_param=user.uid,
                    p_login=False
                )
                
                # Firestoreからデータを読み込み
                try:
                    print(f"Firestoreからデータを読み込み中... ユーザーID: {user.uid}")
                    doc_ref = db.collection('users').document(user.uid)
                    doc = doc_ref.get()
                    if doc.exists:
                        data = doc.to_dict()
                        print(f"Firestoreデータ: {data}")
                        
                        # subjectsフィールドが存在しない場合は、ローカルCSVから読み込み
                        if 'subjects' not in data:
                            print("Firestoreにsubjectsフィールドがありません。ローカルCSVから読み込みます。")
                            ls = load_from_local_csv(user_name)
                            
                            # 読み込んだデータをFirestoreに保存
                            if len(ls) > 1:  # ヘッダー以外にデータがある場合
                                try:
                                    doc_ref.update({'subjects': ls})
                                    print("ローカルデータをFirestoreに保存しました")
                                except Exception as save_error:
                                    print(f"Firestore保存エラー: {save_error}")
                        else:
                            ls = data.get('subjects', [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]])
                        
                        print(f"読み込んだ科目データ: {ls}")
                        print("Firestoreからデータを読み込みました")
                    else:
                        # Firestoreにデータがない場合、ローカルCSVファイルから読み込み
                        print("Firestoreにデータがありません。ローカルCSVファイルから読み込みを試行します。")
                        ls = load_from_local_csv(user_name)
                        
                        # 読み込んだデータをFirestoreに保存
                        if len(ls) > 1:  # ヘッダー以外にデータがある場合
                            try:
                                doc_ref.set({
                                    'email': email,
                                    'user_name': user_name,
                                    'subjects': ls
                                })
                                print("ローカルデータをFirestoreに保存しました")
                            except Exception as save_error:
                                print(f"Firestore保存エラー: {save_error}")
                        else:
                            # データがない場合は初期化
                            ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
                            print("新規ユーザーとしてデータを初期化しました")
                except Exception as db_error:
                    print(f"Firestore読み込みエラー: {db_error}")
                    # エラーの場合もローカルCSVから読み込みを試行
                    ls = load_from_local_csv(user_name)
                
                messagebox.showinfo("成功", f"ログインしました\nユーザー: {user_name}")
                win.close()
                
                # メインウィンドウを再作成し、データを渡す
                from function.gui import GUI
                from function.logic_function import create_table_for_csv, GPA_calc
                
                # テーブルデータを作成
                print(f"テーブル作成前のls: {ls}")
                header, data = create_table_for_csv(state, ls)
                print(f"作成されたヘッダー: {header}")
                print(f"作成されたデータ: {data}")
                
                # メインウィンドウを作成
                main_gui = GUI(state, data, header)
                main_win = main_gui.win
                
                # イベントハンドラーを設定
                from KIT_GPA_calc import setup_event_handlers
                setup_event_handlers(main_win, main_gui, ls, state, auth_service, db)
                
                # GPA計算を実行
                GPA_calc(ls, main_gui)
                
                return main_win, ls
            else:
                messagebox.showerror("エラー", "ユーザーが見つかりません。")
                return win, ls
                
        except Exception as firebase_error:
            print(f"Firebase認証エラー: {firebase_error}")
            messagebox.showerror("エラー", f"Firebase認証に失敗しました: {str(firebase_error)}")
            return win, ls
        
    except Exception as e:
        messagebox.showerror("エラー", f"ログインに失敗しました: {str(e)}")
        return win, ls

def create_user(ls, password, state, win, auth_service, db):
    """Firebase認証を使用した新規ユーザー作成機能"""
    try:
        email = state.current_user_email
        user_name = state.current_user_name
        
        # 基本的なバリデーション
        if not email or not password or not user_name:
            messagebox.showerror("エラー", "すべての項目を入力してください。")
            return win, ls
        
        if "@" not in email:
            messagebox.showerror("エラー", "有効なメールアドレスを入力してください。")
            return win, ls
        
        if len(password) < 6:
            messagebox.showerror("エラー", "パスワードは6文字以上で入力してください。")
            return win, ls
        
        # Firebase認証が利用可能でない場合
        if not auth_service or not db:
            messagebox.showerror("エラー", "Firebase認証サービスが利用できません。\nFirebase設定を確認してください。")
            return win, ls
        
        # Firebase Authenticationを使用してユーザーを作成
        try:
            # Firebase Admin SDKを使用してユーザーを作成
            user = auth_service.create_user(
                email=email,
                password=password,
                display_name=user_name
            )
            
            # ユーザー作成成功
            state.update_state(
                login=True,
                user_name_param=user_name,
                user_email_param=email,
                user_id_param=user.uid,
                p_signup=False
            )
            
            # Firestoreにユーザー情報を保存
            try:
                doc_ref = db.collection('users').document(user.uid)
                doc_ref.set({
                    'email': email,
                    'user_name': user_name,
                    'subjects': [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
                })
                print("Firestoreにユーザー情報を保存しました")
            except Exception as db_error:
                print(f"Firestore保存エラー: {db_error}")
            
            # 初期データを設定
            ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
            
            messagebox.showinfo("成功", f"ユーザーを作成しました\nユーザー: {user_name}")
            win.close()
            
            # メインウィンドウを再作成し、データを渡す
            from function.gui import GUI
            from function.logic_function import create_table_for_csv, GPA_calc
            
            # テーブルデータを作成
            header, data = create_table_for_csv(state, ls)
            
            # メインウィンドウを作成
            main_gui = GUI(state, data, header)
            main_win = main_gui.win
            
            # イベントハンドラーを設定
            from KIT_GPA_calc import setup_event_handlers
            setup_event_handlers(main_win, main_gui, ls, state, auth_service, db)
            
            # GPA計算を実行
            GPA_calc(ls, main_gui)
            
            return main_win, ls
            
        except Exception as firebase_error:
            print(f"Firebase認証エラー: {firebase_error}")
            messagebox.showerror("エラー", f"Firebase認証に失敗しました: {str(firebase_error)}")
            return win, ls
        
    except Exception as e:
        messagebox.showerror("エラー", f"ユーザー作成に失敗しました: {str(e)}")
        return win, ls

def logout_function(state, win, auth_service):
    """ログアウト機能"""
    try:
        # ログアウト処理
        state.update_state(
            login=False,
            user_name_param="",
            user_email_param="",
            user_id_param="",
            idToken_param=""
        )
        
        messagebox.showinfo("成功", "ログアウトしました")
        win.close()
        
        # メインウィンドウを再作成
        from function.gui import GUI
        from function.logic_function import create_table_for_csv, GPA_calc
        
        # 初期データを設定
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
        
        # テーブルデータを作成
        header, data = create_table_for_csv(state, ls)
        
        # メインウィンドウを作成
        main_gui = GUI(state, data, header)
        main_win = main_gui.win
        
        # イベントハンドラーを設定
        from KIT_GPA_calc import setup_event_handlers
        setup_event_handlers(main_win, main_gui, ls, state, auth_service, None)
        
        # GPA計算を実行
        GPA_calc(ls, main_gui)
        
        return main_win, ls
        
    except Exception as e:
        messagebox.showerror("エラー", f"ログアウトに失敗しました: {str(e)}")
        return win, [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]

def delete_user(ls, state, win, auth_service, db):
    """ユーザー削除機能"""
    try:
        # Firebase Admin SDKを使用してユーザーを削除
        if auth_service and state.user_id:
            auth_service.delete_user(state.user_id)
        
        # データベースからユーザーデータを削除
        if db and state.user_id:
            doc_ref = db.collection('users').document(state.user_id)
            doc_ref.delete()
        
        # 状態をリセット
        state.update_state(
            login=False,
            user_name_param="",
            user_email_param="",
            user_id_param="",
            idToken_param=""
        )
        
        messagebox.showinfo("成功", "ユーザーを削除しました")
        win.close()
        
        # メインウィンドウを再作成
        from function.gui import GUI
        from function.logic_function import create_table_for_csv, GPA_calc
        
        # 初期データを設定
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
        
        # テーブルデータを作成
        header, data = create_table_for_csv(state, ls)
        
        # メインウィンドウを作成
        main_gui = GUI(state, data, header)
        main_win = main_gui.win
        
        # イベントハンドラーを設定
        from KIT_GPA_calc import setup_event_handlers
        setup_event_handlers(main_win, main_gui, ls, state, auth_service, None)
        
        # GPA計算を実行
        GPA_calc(ls, main_gui)
        
        return main_win, ls
        
    except Exception as e:
        messagebox.showerror("エラー", f"ユーザー削除に失敗しました: {str(e)}")
        return win, ls

def firebase_save(ls, state, db):
    """Firestoreにデータを保存"""
    try:
        if db and state.isLogin and state.user_id:
            doc_ref = db.collection('users').document(state.user_id)
            doc_ref.update({
                'subjects': ls
            })
            print("データをFirestoreに保存しました")
        return ls
    except Exception as e:
        print(f"データの保存に失敗しました: {e}")
        return ls

def update_subject(ls, subject, units_num, HPT, Pass_Fail, Teacher, state, db):
    """科目情報を更新"""
    try:
        # 科目を検索して更新
        for i, data in enumerate(ls):
            if data[0] == subject:
                ls[i] = [subject, units_num, HPT, Pass_Fail, Teacher]
                break
        
        # Firestoreに保存
        firebase_save(ls, state, db)
        return ls
    except Exception as e:
        print(f"科目の更新に失敗しました: {e}")
        return ls

def delete_subject(ls, subject, state, db):
    """科目を削除"""
    try:
        # 科目を検索して削除
        ls = [data for data in ls if data[0] != subject]
        
        # Firestoreに保存
        firebase_save(ls, state, db)
        return ls
    except Exception as e:
        print(f"科目の削除に失敗しました: {e}")
        return ls

def delete_all_subject(ls, state, db):
    """すべての科目を削除"""
    try:
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
        
        # Firestoreに保存
        firebase_save(ls, state, db)
        return ls
    except Exception as e:
        print(f"全科目の削除に失敗しました: {e}")
        return ls