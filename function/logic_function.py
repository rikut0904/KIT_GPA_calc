import pandas as pd # type: ignore
from function.gui import reload_gui
import webbrowser
from firebase_setting.firebase import login_function, create_user, logout_function, delete_user
import PySimpleGUI as sg

#評価ポイントをローマ字から数字へ変更
def HPT_Checker(HPT):
    HPT = HPT.upper()
    if HPT == "S":
        HPT_num = 4
    elif HPT == "A":
        HPT_num = 3
    elif HPT == "B":
        HPT_num = 2
    elif HPT == "C":
        HPT_num = 1
    elif HPT == "D" or HPT == "F":
        HPT_num = 0
    elif HPT == "合" or HPT == "否":
        HPT_num = ""
    else:
        HPT_num = "error"
    print("HPT:",HPT)
    print("HPT_num:",HPT_num)
    return HPT_num

#科目名の重複チェック
def check_subject(ls, subject):
    if len(ls) > 1:
        for data in ls[1:]:
            if data[0] != subject:
                return True
        return False    
    else:
        return True

#科目のエラーチェック
def check_subject_error(ls, subject, units_num, HPT, Pass_Fail):
    HPT_num = HPT_Checker(HPT)
    if HPT_num != "error":
        if check_subject(ls, subject):
            ls.append([subject, units_num, HPT, Pass_Fail])
            txt = ""
        else:
            txt = "subject error"
    else:
        txt = "input error" if HPT_num != "error" else "Point input error"
    return txt, ls

#GPAの計算
def GPA_calc(ls):
    total_HPT, total_units_num, all_total_units_num = 0, 0, 0
    txt = ""
    try:
        if len(ls) > 1:
            for data in ls[1:]:
                subject, units_num, HPT, Pass_Fail = data
                HPT_num = HPT_Checker(HPT)
                if HPT_num == "":
                    if HPT == "合":
                        print("合格")
                        all_total_units_num += units_num
                    else:
                        print("不合格")
                elif HPT_num != "error":
                    if HPT_num > 0:
                        print("not error:",HPT_num)
                        total_HPT += HPT_num * units_num
                        total_units_num += units_num
                        all_total_units_num += units_num
                    else:
                        print("落単科目")
                else:
                    txt = "input error"
        else:
            txt = "入力がありません"
    except Exception as e:
        print(f"GPA計算に失敗しました: {e}")
    return total_HPT, total_units_num, all_total_units_num, txt

#CSVファイルを読み取り、表を作成
def create_table_for_csv(state):
    df = pd.read_csv(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv")
    data = df.values.tolist()
    header_list = list(df.columns)
    return header_list, data

def setting_function(ls, state, win, eve, val, auth, db):
    if eve == "-setting_exit-":
        state.update_state(p_setting=False)
        win = reload_gui(state, win)
    elif eve == "-Auth_exit-":
        state.update_state(p_login=False, p_signup=False)
        win = reload_gui(state, win)
    elif eve == "-back_Auth-":
        state.update_state(p_login=True, p_signup=False)
        win = reload_gui(state, win)
    elif eve == "-GitHub-":
        webbrowser.open("https://github.com/rikut0904/KIT_GPA_calc")
    elif eve == "-Help-":
        webbrowser.open("https://github.com/rikut0904/KIT_GPA_calc/blob/main/README.md")
    elif eve == "-BugReport-":
        webbrowser.open("https://docs.google.com/forms/d/e/1FAIpQLSePf3v6STDK2kt503UBAxdxW_2DFT0y9UE7Fh9GdQaSIKtp2w/viewform?usp=header")
    elif eve == "-Update-":
        webbrowser.open("https://shrouded-rain-eb7.notion.site/KIT_GPA_calc-1a4b6247b0898072896de3a2fa534dbf?pvs=4")
    elif eve == "-Contact-":
        webbrowser.open("https://docs.google.com/forms/d/e/1FAIpQLSdEor5q8YYRHpPszyifnrUC0lp4JwP0_t8t-zRIuQDy4CtM2Q/viewform?usp=header")
    elif eve == "-Login-":  #ログインボタンが押された際の動作
        state.update_state(p_login=True)
        win = reload_gui(state, win)
    elif eve == "-Auth-":   #ログイン
        email = val["-email-"]
        password = val["-password-"]
        state.update_state(user_email_param=email)
        win, ls = login_function(ls, password, state, win, auth, db)
    elif eve == "-UserCreate-":  #ユーザー作成ボタンが押された際の動作
        state.update_state(p_signup=True)
        win = reload_gui(state, win)
    elif eve == "-Auth_Create-":  #新規登録
        email = val["-email-"]
        UserName = val["-UserName-"]
        password = val["-password-"]
        state.update_state(user_name_param=UserName, user_email_param=email)
        win, ls = create_user(ls, password, state, win, auth, db)
    elif eve == "-Logout-": # ログアウト
        ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
        print(ls)
        win = logout_function(state, win, auth, db)
        win["-UserName-"].update(state.user_name)
        win["-email-"].update(state.user_email)
    elif eve == "-Delete-":
        delete_eve = sg.popup_yes_no("ユーザーを削除しますか？", font = (None, 15))
        if delete_eve == "Yes":
            ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
            print(ls)
            win, ls = delete_user(ls, state, win, auth, db)
    return win, ls
