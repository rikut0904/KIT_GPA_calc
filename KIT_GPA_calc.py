import csv
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from firebase_setting.firebase import get_auth, get_database, firebase_save, update_subject, delete_subject, delete_all_subject, initialize_firebase
from function import state
from function.gui import GUI, reload_gui
from function.logic_function import GPA_calc, create_table_for_csv, setting_function, check_subject_error

# 必要なサービスを取得
# Firebase初期化
if initialize_firebase():
    auth = get_auth()
    db = get_database()
else:
    auth = None
    db = None
    print("Firebase初期化に失敗しました。オフラインモードで動作します。")

def setup_event_handlers(win, gui, ls, state, auth, db):
    """イベントハンドラーを設定する関数"""
    
    def on_submit(event):
        nonlocal ls  # 外側のls変数を参照
        val = gui.get_values()
        if val["-subject-"] != "" and val["-units_num-"] != "" and val["-HPT-"] != "":
            subject = val["-subject-"]
            units_num = int(val["-units_num-"])
            HPT = val["-HPT-"]
            Pass_Fail = val["-Pass/Fail-"]
            Teacher = val["-Teacher-"]
            txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail, Teacher)
            if state.isLogin and auth and db:
                ls = firebase_save(ls, state, db)
            print(ls)
            gui.update_widget("-txt-", txt)
            gui.update_widget("-subject-", "")
            gui.update_widget("-units_num-", "")
            gui.update_widget("-HPT-", "")
            gui.update_widget("-Pass/Fail-", False)
            gui.update_widget("-Teacher-", False)
        else:
            txt = "必要事項を入力してください。"
            gui.update_widget("-txt-", txt)
    
    def on_final(event):
        GPA_calc(ls, gui)
    
    def on_file_import(event):
        nonlocal ls  # 外側のls変数を参照
        val = gui.get_values()
        File_name = val["-inputFilePath-"]
        if File_name:
            try:
                with open(File_name, newline="", encoding="utf-8") as old_GPA_data:
                    file_csv = csv.reader(old_GPA_data)
                    next(file_csv)  # ヘッダー行をスキップ
                    
                    imported_count = 0
                    skipped_count = 0
                    for row in file_csv:
                        if len(row) >= 5:  # 必要な列数があるかチェック
                            subject = row[0]
                            units_num = int(row[1])
                            HPT = row[2]
                            # 文字列の"True"/"False"をbool型に変換
                            Pass_Fail = row[3].lower() == 'true'
                            Teacher = row[4].lower() == 'true'
                            
                            txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail, Teacher)
                            if txt == "":  # エラーがない場合
                                imported_count += 1
                            elif txt == "subject error":  # 重複の場合
                                skipped_count += 1
                                print(f"科目 '{subject}' は既に存在するためスキップしました。")
                            else:
                                print(f"科目 '{subject}' のインポートでエラー: {txt}")
                    
                    if imported_count > 0:
                        txt = f"ファイルが正常にインポートされました。{imported_count}件の科目を追加しました。"
                        if skipped_count > 0:
                            txt += f" {skipped_count}件の重複科目をスキップしました。"
                    else:
                        txt = "インポートできる科目がありませんでした。"
                        if skipped_count > 0:
                            txt += f" {skipped_count}件の重複科目をスキップしました。"
                        
                if state.isLogin and auth and db:
                    ls = firebase_save(ls, state, db)
                print(txt)
            except Exception as e:
                txt = f"ファイルのインポートに失敗しました。：{str(e)}"
                print(f"詳細エラー: {e}")
        else:
            txt = "ファイルが選択されていません。"
        gui.update_widget("-txt-", txt)
        gui.update_widget("-inputFilePath-", "")
        GPA_calc(ls, gui)
    
    def on_csv(event):
        with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(ls)
        header, data = create_table_for_csv(state)
        show_csv_window(data, header)
    
    def on_subject_ui(event):
        with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(ls)
        header, data = create_table_for_csv(state)
        state.update_state(p_subject_delete=True)
        subject_edit_window = GUI(state, data, header)
        setup_subject_edit_handlers(subject_edit_window, ls, state, db)
    
    def on_gpa_reset(event):
        nonlocal ls  # 外側のls変数を参照
        res = messagebox.askyesno("確認", "成績情報をリセットしますか？\n※ログイン中はデータベースが削除されます。")
        if res:
            if state.isLogin and auth and db:
                ls = delete_all_subject(ls, state, db)
            else:
                ls.clear()
            ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
            GPA = 0.00
            SGPT = 0.00
            txt = ""
            gui.update_widget("-GPA-", f"{GPA:.2f}")
            gui.update_widget("-all_total_units_num-", f"{0}単位")
            gui.update_widget("-Graduation_total_units_num-", f"{0}単位")
            gui.update_widget("-SGPT-", f"{SGPT:.2f}")
            gui.update_widget("-txt-", txt)
        else:
            messagebox.showinfo("情報", "成績情報をリセットしませんでした。")
    
    def on_setting(event):
        state.update_state(p_setting=True)
        setting_window = GUI(state)
        setup_setting_handlers(setting_window, ls, state, auth, db)
    
    # イベントをバインド
    win.bind("<<Submit>>", on_submit)
    win.bind("<<Final>>", on_final)
    win.bind("<<File_Import>>", on_file_import)
    win.bind("<<CSV>>", on_csv)
    win.bind("<<Subject_UI>>", on_subject_ui)
    win.bind("<<GPA_reset>>", on_gpa_reset)
    win.bind("<<Setting>>", on_setting)

def show_csv_window(data, header):
    """CSVファイル内容を表示するウィンドウ"""
    csv_window = tk.Toplevel()
    csv_window.title("CSVファイル内容")
    csv_window.geometry("700x400")
    csv_window.grab_set()
    
    # テーブル表示
    tree = ttk.Treeview(csv_window, columns=header, show='headings', height=20)
    
    for col in header:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    
    # スクロールバー
    scrollbar = ttk.Scrollbar(csv_window, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=10)
    
    # データを挿入
    for row in data:
        tree.insert('', tk.END, values=row)
    
    # 閉じるボタン
    ttk.Button(csv_window, text="Close", command=csv_window.destroy).pack(pady=10)

def setup_subject_edit_handlers(subject_edit_window, ls, state, db):
    """科目修正・削除ウィンドウのイベントハンドラーを設定"""
    
    def on_subject_update(event):
        nonlocal ls  # 外側のls変数を参照
        val = subject_edit_window.get_values()
        subject = val["-subject-"]
        units_num = int(val["-units_num-"]) if val["-units_num-"] != "" else ""
        HPT = val["-HPT-"] if val["-HPT-"] != "" else ""
        Pass_Fail = val["-Pass/Fail-"] if val["-Pass/Fail-"] != "" else ""
        Teacher = val["-Teacher-"] if val["-Teacher-"] != "" else ""
        print(subject, units_num, HPT, Pass_Fail, Teacher)
        ls = update_subject(ls, subject, units_num, HPT, Pass_Fail, Teacher, state, db)
        with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(ls)
        header, data = create_table_for_csv(state)
        subject_edit_window.close()
        new_subject_edit_window = GUI(state, data, header)
        setup_subject_edit_handlers(new_subject_edit_window, ls, state, db)
    
    def on_subject_delete(event):
        nonlocal ls  # 外側のls変数を参照
        val = subject_edit_window.get_values()
        subject = val["-subject-"]
        ls = delete_subject(ls, subject, state, db)
        with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(ls)
        header, data = create_table_for_csv(state)
        subject_edit_window.close()
        new_subject_edit_window = GUI(state, data, header)
        setup_subject_edit_handlers(new_subject_edit_window, ls, state, db)
    
    def on_close(event):
        state.update_state(p_subject_delete=False)
        subject_edit_window.close()
    
    subject_edit_window.win.bind("<<Subject_update>>", on_subject_update)
    subject_edit_window.win.bind("<<Subject_delete>>", on_subject_delete)
    subject_edit_window.win.bind("<<Close>>", on_close)

def setup_setting_handlers(setting_window, ls, state, auth, db):
    """設定ウィンドウのイベントハンドラーを設定"""
    
    def on_setting_exit(event):
        state.update_state(p_setting=False)
        setting_window.close()
    
    def on_auth_exit(event):
        state.update_state(p_login=False, p_signup=False)
        setting_window.close()
    
    def on_back_auth(event):
        state.update_state(p_login=True, p_signup=False)
        setting_window.close()
        login_window = GUI(state)
        setup_login_handlers(login_window, ls, state, auth, db)
    
    def on_github(event):
        import webbrowser
        webbrowser.open("https://github.com/rikut0904/KIT_GPA_calc")
    
    def on_help(event):
        import webbrowser
        webbrowser.open("https://github.com/rikut0904/KIT_GPA_calc/blob/main/README.md")
    
    def on_bug_report(event):
        import webbrowser
        webbrowser.open("https://docs.google.com/forms/d/e/1FAIpQLSePf3v6STDK2kt503UBAxdxW_2DFT0y9UE7Fh9GdQaSIKtp2w/viewform?usp=header")
    
    def on_update(event):
        import webbrowser
        webbrowser.open("https://shrouded-rain-eb7.notion.site/KIT_GPA_calc-1a4b6247b0898072896de3a2fa534dbf?pvs=4")
    
    def on_contact(event):
        import webbrowser
        webbrowser.open("https://docs.google.com/forms/d/e/1FAIpQLSdEor5q8YYRHpPszyifnrUC0lp4JwP0_t8t-zRIuQDy4CtM2Q/viewform?usp=header")
    
    def on_login(event):
        state.update_state(p_login=True)
        setting_window.close()
        login_window = GUI(state)
        setup_login_handlers(login_window, ls, state, auth, db)
    
    def on_auth(event):
        nonlocal ls  # 外側のls変数を参照
        val = setting_window.get_values()
        email = val["-email-"]
        password = val["-password-"]
        state.update_state(user_email_param=email)
        from firebase_setting.firebase import login_function
        login_window, ls = login_function(ls, password, state, setting_window, auth, db)
    
    def on_user_create(event):
        state.update_state(p_signup=True)
        setting_window.close()
        signup_window = GUI(state)
        setup_signup_handlers(signup_window, ls, state, auth, db)
    
    def on_auth_create(event):
        nonlocal ls  # 外側のls変数を参照
        val = setting_window.get_values()
        email = val["-email-"]
        UserName = val["-UserName-"]
        password = val["-password-"]
        state.update_state(user_name_param=UserName, user_email_param=email)
        from firebase_setting.firebase import create_user
        signup_window, ls = create_user(ls, password, state, setting_window, auth, db)
    
    def on_logout(event):
        nonlocal ls  # 外側のls変数を参照
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
        print(ls)
        from firebase_setting.firebase import logout_function
        new_window, ls = logout_function(state, setting_window, auth)
        if new_window:
            new_window.update_widget("-UserName-", state.user_name)
            new_window.update_widget("-email-", state.user_email)
    
    def on_delete(event):
        nonlocal ls  # 外側のls変数を参照
        res = messagebox.askyesno("確認", "ユーザーを削除しますか？")
        if res:
            ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
            print(ls)
            from firebase_setting.firebase import delete_user
            setting_window, ls = delete_user(ls, state, setting_window, auth, db)
    
    # イベントをバインド
    setting_window.win.bind("<<setting_exit>>", on_setting_exit)
    setting_window.win.bind("<<Auth_exit>>", on_auth_exit)
    setting_window.win.bind("<<back_Auth>>", on_back_auth)
    setting_window.win.bind("<<GitHub>>", on_github)
    setting_window.win.bind("<<Help>>", on_help)
    setting_window.win.bind("<<BugReport>>", on_bug_report)
    setting_window.win.bind("<<Update>>", on_update)
    setting_window.win.bind("<<Contact>>", on_contact)
    setting_window.win.bind("<<Login>>", on_login)
    setting_window.win.bind("<<Auth>>", on_auth)
    setting_window.win.bind("<<UserCreate>>", on_user_create)
    setting_window.win.bind("<<Auth_Create>>", on_auth_create)
    setting_window.win.bind("<<Logout>>", on_logout)
    setting_window.win.bind("<<Delete>>", on_delete)

def setup_login_handlers(login_window, ls, state, auth, db):
    """ログインウィンドウのイベントハンドラーを設定"""
    
    def on_auth(event):
        nonlocal ls, login_window  # 外側のls変数とlogin_window変数を参照
        val = login_window.get_values()
        email = val["-email-"]
        password = val["-password-"]
        state.update_state(user_email_param=email)
        from firebase_setting.firebase import login_function
        login_window, ls = login_function(ls, password, state, login_window, auth, db)
    
    def on_user_create(event):
        state.update_state(p_signup=True)
        login_window.close()
        signup_window = GUI(state)
        setup_signup_handlers(signup_window, ls, state, auth, db)
    
    def on_auth_exit(event):
        state.update_state(p_login=False, p_signup=False)
        login_window.close()
    
    login_window.win.bind("<<Auth>>", on_auth)
    login_window.win.bind("<<UserCreate>>", on_user_create)
    login_window.win.bind("<<Auth_exit>>", on_auth_exit)

def setup_signup_handlers(signup_window, ls, state, auth, db):
    """新規登録ウィンドウのイベントハンドラーを設定"""
    
    def on_auth_create(event):
        nonlocal ls, signup_window  # 外側のls変数とsignup_window変数を参照
        val = signup_window.get_values()
        email = val["-email-"]
        UserName = val["-UserName-"]
        password = val["-password-"]
        state.update_state(user_name_param=UserName, user_email_param=email)
        from firebase_setting.firebase import create_user
        signup_window, ls = create_user(ls, password, state, signup_window, auth, db)
    
    def on_back_auth(event):
        state.update_state(p_login=True, p_signup=False)
        signup_window.close()
        login_window = GUI(state)
        setup_login_handlers(login_window, ls, state, auth, db)
    
    signup_window.win.bind("<<Auth_Create>>", on_auth_create)
    signup_window.win.bind("<<back_Auth>>", on_back_auth)

#メインプログラム
def main():
    ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
    state.update_state(login=False, p_login=False, p_signup=False, p_setting=False)
    gui = GUI(state)
    win = gui.win
    
    # イベントハンドラーを設定
    setup_event_handlers(win, gui, ls, state, auth, db)
    
    # メインループを開始
    win.mainloop()

if __name__ == "__main__":
    main()