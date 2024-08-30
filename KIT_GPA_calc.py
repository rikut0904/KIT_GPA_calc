import csv
import pandas as pd # type: ignore
import PySimpleGUI as sg
#GUIデザイン
sg.theme("Default1")
class GUI():
    def __init__(self):
        #GUIレイアウトの設定
        layout = [[sg.T("　　　　  累積GPA："),sg.T("0.0", key="-GPA-")],
                  [sg.T("正課学習ポイント："),sg.T("0.0", key="-SGPT-")],
                  [sg.T("", key = "-txt-")],
                  [sg.T("　　　 科目名："),sg.I("", key="-subject-",expand_x=True)],
                  [sg.T("　　　 単位数："),sg.I("", key="-units_num-",expand_x=True)],
                  [sg.T("評価ポイント："),sg.I("", key="-HPT-",expand_x=True)],
                  [sg.Checkbox("合否科目",default=False, key="-Pass/Fail-")],
                  [sg.T("過去のsubject_grades_data.csvファイル："),sg.I("", key="-inputFilePath-", expand_x=True),
                   sg.FileBrowse("ファイル選択"),],
                  [sg.Button("Submit", key="-Btn-"), sg.Button("Final", key="-Final-"),
                   sg.Button("CSVファイル", key="-CSV-"), sg.Button("ファイルインポート", key="-File_Import-"),
                   sg.Button("GPAリセット", key="-GPA_reset-")]]
        self.win = sg.Window("GPA計算", layout, font = (None,15),
                             finalize=True, resizable = True)


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


#GPAの計算
def GPA_calc(ls,subject,units_num,HPT,Pass_Fail,total_HPT,total_units_num, all_total_units_num):
    HPT_num = HPT_Checker(HPT)
    if HPT_num == "":
        total_HPT, total_units_num = total_HPT, total_units_num
        all_total_units_num += units_num
        ls.append([subject, units_num, HPT, Pass_Fail])
    elif HPT_num != "error":
        if not Pass_Fail:
            total_HPT += HPT_num * units_num
            total_units_num += units_num
        all_total_units_num += units_num
        ls.append([subject, units_num, HPT, Pass_Fail])
    return ls,HPT_num, total_HPT, total_units_num, all_total_units_num


#CSVファイルを読み取り、表を作成
def create_table_for_csv():
    df = pd.read_csv("subject_grades_data.csv")
    data = df.values.tolist()
    header_list = list(df.columns)
    return header_list, data


#メインプログラム
def main():
    ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
    total_HPT, total_units_num, all_total_units_num = 0, 0, 0
    gui = GUI()
    win = gui.win
    while True:
        eve, val = win.read()
        if eve == sg.WIN_CLOSED:
            break
        elif eve == "-Btn-":    #Submitボタンが押された際の動作
            #科目名、単位数、評価ポイントをGUIより入力しリストに格納
            if val["-subject-"] != "" and val["-units_num-"] != "" and val["-HPT-"] != "":
                subject = val["-subject-"]
                units_num = int(val["-units_num-"])
                HPT = val["-HPT-"]
                Pass_Fail = val["-Pass/Fail-"]
                ls, HPT_num, total_HPT, total_units_num, all_total_units_num = GPA_calc(ls, subject, units_num, HPT, Pass_Fail, total_HPT, total_units_num, all_total_units_num)
                txt = "" if HPT_num != "error" else "Point input error"
                win["-txt-"].update(txt)
                win["-subject-"].update("")
                win["-units_num-"].update("")
                win["-HPT-"].update("")
            else:
                txt = "必要事項を入力してください。"
                win["-txt-"].update(txt)
        elif eve == "-Final-":  #finalボタンが押された際の動作
            #Submitで入力された成績情報をもとにGPAを計算
            if total_HPT != 0 or total_units_num != 0:
                GPA = total_HPT / total_units_num
                SGPT = GPA * all_total_units_num
                win["-GPA-"].update(f'{GPA:.1f}')
                win["-SGPT-"].update(f'{SGPT:.1f}')
            else:
                txt = "成績を入力またはCSVファイルをインポートしてください。"
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
                            Pass_Fail = val["-Pass/Fail-"]
                            ls, HPT_num, total_HPT, total_units_num, all_total_units_num = GPA_calc(ls, subject, units_num, HPT, Pass_Fail, total_HPT, total_units_num, all_total_units_num)
                    txt = ("ファイルが正常にインポートされました。")
                except Exception as e:
                    txt = f"ファイルのインポートに失敗しました。：{str(e)}"
            else:
                txt = "ファイルが選択されていません。"
            GPA = 0.0
            SGPT = 0.0
            win["-GPA-"].update(GPA)
            win["-SGPT-"].update(SGPT)
            win["-inputFilePath-"].update("")
            win["-txt-"].update(txt)
        elif eve == "-CSV-":    #CSVファイルボタンが押された際の動作
            #Submitで入力された成績情報をCSVファイル化させ表形式でGUIに表示
            with open("subject_grades_data.csv", "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(ls)
            sg.popup_quick("CSVファイルが保存されました。")
            header, data = create_table_for_csv()
            layout_table = [[sg.Table(values=data, headings=header, display_row_numbers=True,
                                      auto_size_columns=True, num_rows=min(25, len(data)),
                                      expand_x=True, expand_y=True)],
                            [sg.Button("Close", key = "-Close-")]]
            table_window = sg.Window('CSVファイル内容', layout_table, font = (None,15),
                                     size=(700,150), finalize=True, resizable = True)
            while True:
                event, value = table_window.read()
                if event == sg.WIN_CLOSED or event == "-Close-":
                    break
            table_window.close()
        elif eve == "-GPA_reset-":  #GPAリセットボタンが押された際の動作
            #成績情報を削除してよいかを確認し、成績情報を削除する
            res = sg.popup_yes_no("成績情報をリセットしますか？")
            if res == "Yes":
                ls.clear()
                ls = [["科目名", "単位数", "評価ポイント", "合否科目"]]
                total_HPT, total_units_num, all_total_units_num = 0, 0 ,0
                GPA = 0.0
                SGPT = 0.0
                txt =""
                win["-GPA-"].update(GPA)
                win["-SGPT-"].update(SGPT)
                win["-txt-"].update(txt)
            else:
                sg.popup_quick_message("成績情報をリセットしませんでした。")
    win.close()


if __name__ == "__main__":
    main()