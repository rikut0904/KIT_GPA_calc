import csv
import pandas as pd
import PySimpleGUI as sg
from firebase_setting.firebase import get_auth, get_database, firebase_save, update_subject, delete_subject, delete_all_subject
from function import state
from function.gui import GUI, reload_gui
from function.logic_function import GPA_calc, create_table_for_csv, setting_function, check_subject_error

# Todo: ログ、タイムスタンプの調整

# 必要なサービスを取得
auth = get_auth()
db = get_database()

#メインプログラム
def main():
    ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
    state.update_state(login=False, p_login=False, p_signup=False, p_setting=False)
    gui = GUI(state)
    win = gui.win
    while True:
        print(ls)
        eve, val = win.read()
        if eve == sg.WIN_CLOSED:
            break
        elif eve == "-Setting-":
            state.update_state(p_setting=True)
            win = reload_gui(state, win)
            win["-UserName-"].update(state.user_name)
            win["-email-"].update(state.user_email)
        elif state.popup_setting:
            win, ls = setting_function(ls, state, win, eve, val, auth, db)
        elif eve == "-Submit-":    #Submitボタンが押された際の動作
            #科目名、単位数、評価ポイントをGUIより入力しリストに格納
            if val["-subject-"] != "" and val["-units_num-"] != "" and val["-HPT-"] != "":
                subject = val["-subject-"]
                units_num = int(val["-units_num-"])
                HPT = val["-HPT-"]
                Pass_Fail = val["-Pass/Fail-"]
                Teacher = val["-Teacher-"]
                txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail, Teacher)
                if state.isLogin:
                    ls = firebase_save(ls, state, db)
                print(ls)
                win["-txt-"].update(txt)
                win["-subject-"].update("")
                win["-units_num-"].update("")
                win["-HPT-"].update("")
                win["-Pass/Fail-"].update(False)
                win["-Teacher-"].update(False)
            else:
                txt = "必要事項を入力してください。"
                win["-txt-"].update(txt)
        elif eve == "-Final-":  #finalボタンが押された際の動作
            #Submitで入力された成績情報をもとにGPAを計算
            GPA_calc(ls, win)
        elif eve == "-File_Import-":    #ファイルインポートボタンが押された際の動作
            #外部からCSVファイルをインポートし成績情報を入力する。
            File_name = val["-inputFilePath-"]
            if File_name:
                try:
                    with open(File_name, newline="", encoding="utf-8") as old_GPA_data:
                        file_csv = csv.reader(old_GPA_data)
                        next(file_csv)
                        for row in file_csv:
                            subject = row[0]
                            units_num = int(row[1])
                            HPT = row[2]
                            Pass_Fail = row[3]
                            Teacher = row[4]
                            txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail, Teacher)
                            txt = ("ファイルが正常にインポートされました。")
                    if state.isLogin:
                        ls = firebase_save(ls, state, db)
                    print(txt)
                except Exception as e:
                    txt = f"ファイルのインポートに失敗しました。：{str(e)}"
            else:
                txt = "ファイルが選択されていません。"
            win["-txt-"].update(txt)
            win["-inputFilePath-"].update("")
            GPA_calc(ls, win)
        elif eve == "-CSV-" or eve == "-Subject_UI-":    #CSVファイルボタンが押された際の動作
            #Submitで入力された成績情報をCSVファイル化させ表形式でGUIに表示
            with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(ls)
            header, data = create_table_for_csv(state)
            if eve == "-CSV-":
                layout_tb = [[sg.Table(values=data, headings=header, display_row_numbers=True,
                                        auto_size_columns=True, num_rows=min(25, len(data)),
                                        expand_x=True, expand_y=True)],
                                [sg.Button("Close", key = "-Close-")]]
                tb_win = sg.Window('CSVファイル内容', layout_tb, font = (None,15),
                                        size=(700,150), finalize=True, resizable = True)
                while True:
                    tb_eve, tb_val = tb_win.read()
                    if tb_eve == sg.WIN_CLOSED or tb_eve == "-Close-":
                        break
                tb_win.close()
            elif eve == "-Subject_UI-":
                state.update_state(p_subject_delete=True)
                win = reload_gui(state, win, data, header)
                while True:
                    eve, val = win.read()
                    if eve == sg.WIN_CLOSED or eve == "-Close-":
                        break
                    elif eve == "-Subject_update-":
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
                        win = reload_gui(state, win, data, header)
                    elif eve == "-Subject_delete-":
                        subject = val["-subject-"]
                        ls = delete_subject(ls, subject, state, db)
                        with open(f"subject_grades_data_{'Guest' if state.user_name == '' else state.user_name}.csv", "w", newline="", encoding="utf-8") as f:
                            writer = csv.writer(f)
                            writer.writerows(ls)
                        header, data = create_table_for_csv(state)  
                        win = reload_gui(state, win, data, header)
                state.update_state(p_subject_delete=False)
                win = reload_gui(state, win)
        elif eve == "-GPA_reset-":  #GPAリセットボタンが押された際の動作
            #成績情報を削除してよいかを確認し、成績情報を削除する
            res = sg.popup_yes_no("成績情報をリセットしますか？\n※ログイン中はデータベースが削除されます。")
            if res == "Yes":
                if state.isLogin:
                    ls = delete_all_subject(ls, state, db)
                else:
                    ls.clear()
                ls = [["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]]
                GPA = 0.00
                SGPT = 0.00
                txt =""
                win["-GPA-"].update(GPA)
                win["-all_total_units_num-"].update(f"{0}単位")
                win["-Graduation_total_units_num-"].update(f"{0}単位")
                win["-SGPT-"].update(SGPT)
                win["-txt-"].update(txt)
            else:
                sg.popup_quick_message("成績情報をリセットしませんでした。")
    win.close()


if __name__ == "__main__":
    main()