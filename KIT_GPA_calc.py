import csv
import pandas as pd
import PySimpleGUI as sg
from firebase_setting.firebase import get_auth, get_database, firebase_save, delete_subject
from function import state
from function.gui import GUI, reload_gui
from function.logic_function import GPA_calc, create_table_for_csv, setting_function, check_subject_error

# Todo: 成績情報を変更する

# 必要なサービスを取得
auth = get_auth()
db = get_database()

#メインプログラム
def main():
    ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
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
                txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail)
                if state.isLogin:
                    ls = firebase_save(ls, state, db)
                print(ls)
                win["-txt-"].update(txt)
                win["-subject-"].update("")
                win["-units_num-"].update("")
                win["-HPT-"].update("")
                win["-Pass/Fail-"].update(False)
            else:
                txt = "必要事項を入力してください。"
                win["-txt-"].update(txt)
        elif eve == "-Final-":  #finalボタンが押された際の動作
            #Submitで入力された成績情報をもとにGPAを計算
            total_HPT, total_units_num, all_total_units_num, txt = GPA_calc(ls)
            if total_HPT != 0 or total_units_num != 0:
                print(ls)
                GPA = total_HPT / total_units_num
                SGPT = GPA * all_total_units_num
                win["-GPA-"].update(f'{GPA:.2f}')
                win["-SGPT-"].update(f'{SGPT:.2f}')
                win["-all_total_units_num-"].update(f"{all_total_units_num}単位")
                win["-txt-"].update(txt)
            else:
                win["-txt-"].update(txt)
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
                            txt, ls = check_subject_error(ls, subject, units_num, HPT, Pass_Fail)
                            if state.isLogin:
                                ls = firebase_save(ls, state, db)
                            txt = ("ファイルが正常にインポートされました。")
                    print(ls)
                except Exception as e:
                    txt = f"ファイルのインポートに失敗しました。：{str(e)}"
            else:
                txt = "ファイルが選択されていません。"
            GPA = 0.00
            SGPT = 0.00
            win["-GPA-"].update(GPA)
            win["-SGPT-"].update(SGPT)
            win["-inputFilePath-"].update("")
            win["-txt-"].update(txt)
        elif eve == "-CSV-" or eve == "-Subject_Delete_UI-":    #CSVファイルボタンが押された際の動作
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
            elif eve == "-Subject_Delete_UI-":
                state.update_state(p_subject_delete=True)
                win = reload_gui(state, win, data, header)
                while True:
                    eve, val = win.read()
                    if eve == sg.WIN_CLOSED or eve == "-Close-":
                        break
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
            res = sg.popup_yes_no("成績情報をリセットしますか？\n※データベースは削除されません。")
            if res == "Yes":
                ls.clear()
                ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
                total_HPT, total_units_num, all_total_units_num = 0, 0 ,0
                GPA = 0.00
                SGPT = 0.00
                txt =""
                win["-GPA-"].update(GPA)
                win["-all_total_units_num-"].update(f"{0}単位")
                win["-SGPT-"].update(SGPT)
                win["-txt-"].update(txt)
            else:
                sg.popup_quick_message("成績情報をリセットしませんでした。")
    win.close()


if __name__ == "__main__":
    main()