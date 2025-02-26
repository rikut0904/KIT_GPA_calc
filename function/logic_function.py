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
def check_subject_error(ls, subject, units_num, HPT, Pass_Fail, Teacher):
    HPT_num = HPT_Checker(HPT)
    if HPT_num != "error":
        if check_subject(ls, subject):
            ls.append([subject, units_num, HPT, Pass_Fail, Teacher])
            txt = ""
        else:
            txt = "subject error"
    else:
        txt = "input error" if HPT_num != "error" else "Point input error"
    return txt, ls

#GPAの計算
def GPA_calc(ls, win, setting=False):
    total_HPT, total_units_num, Graduation_total_units_num, all_total_units_num = 0, 0, 0, 0
    GPA, SGPT = 0.00, 0.00
    txt = ""
    try:
        if len(ls) > 1:
            for data in ls[1:]:
                subject, units_num, HPT, Pass_Fail, Teacher = data
                HPT_num = HPT_Checker(HPT)
                if Teacher:
                    print(f"教職科目:{subject} {units_num}単位")
                    all_total_units_num += units_num
                else:
                    if HPT_num == "":
                        if HPT == "合":
                            print(f"合格:{subject} {units_num}単位")
                            Graduation_total_units_num += units_num
                        else:
                            print(f"不合格:{subject} {units_num}単位")
                    elif HPT_num != "error":
                        if HPT_num > 0:
                            print(f"not error:{subject} {units_num}単位")
                            total_HPT += HPT_num * units_num
                            total_units_num += units_num
                        elif HPT_num == 0:
                            print(f"落単科目:{subject} {units_num}単位")
                    else:
                        txt = f"input error:{subject} {units_num}単位"
            print(f"全評価点数:{total_HPT}, GPA計算科目数:{total_units_num}, 卒業科目数:{Graduation_total_units_num}, 全科目数:{all_total_units_num}")
            Graduation_total_units_num += total_units_num
            print(f"全評価点数:{total_HPT}, GPA計算科目数:{total_units_num}, 卒業科目数:{Graduation_total_units_num}, 全科目数:{all_total_units_num}")
            all_total_units_num += Graduation_total_units_num
        else:
            txt = "入力がありません"

    except Exception as e:
        print(f"GPA計算に失敗しました: {e}")
    print(f"全評価点数:{total_HPT}, GPA計算科目数:{total_units_num}, 卒業科目数:{Graduation_total_units_num}, 全科目数:{all_total_units_num}")
    if total_HPT != 0 or total_units_num != 0:
        GPA = total_HPT / total_units_num
        SGPT = GPA * Graduation_total_units_num
        print(ls)
        win["-GPA-"].update(f'{GPA:.2f}')
        win["-SGPT-"].update(f'{SGPT:.2f}')
        win["-Graduation_total_units_num-"].update(f"{Graduation_total_units_num}単位")
        win["-all_total_units_num-"].update(f"{all_total_units_num}単位")
        win["-txt-"].update(txt)
    else:
        if setting and txt == "入力がありません":
            txt = ""
        win["-txt-"].update(txt)

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
        GPA_calc(ls, win, setting=True)
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
        ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
        print(ls)
        win, ls  = logout_function(state, win, auth)
        win["-UserName-"].update(state.user_name)
        win["-email-"].update(state.user_email)
    elif eve == "-Delete-":
        delete_eve = sg.popup_yes_no("ユーザーを削除しますか？", font = (None, 15))
        if delete_eve == "Yes":
            ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
            print(ls)
            win, ls = delete_user(ls, state, win, auth, db)
    return win, ls
