import PySimpleGUI as sg

#GUIデザイン
sg.theme("Default1")
class GUI():
    def __init__(self,state, data=None, header=None):
        if state.popup_login: #ログイン画面
            layout = [[sg.T("新規登録") if state.popup_signup else sg.T("ログイン")],
                      [sg.T("", key = "-txt-")],
                      [sg.T("メールアドレス："),sg.I("", key="-email-",expand_x=True)],
                      [(sg.T("ユーザー名    ："),sg.I("", key="-UserName-",expand_x=True))
                       if state.popup_signup else sg.T("")],
                      [sg.T("パスワード　   ："),sg.I("", key="-password-",expand_x=True)],
                      [(sg.Button("登録", key="-Auth_Create-"), sg.Button("戻る",key="-back_Auth-"))
                       if state.popup_signup else (sg.Button("ログイン", key="-Auth-"),
                                             sg.Button("新規登録", key="-UserCreate-"),
                                             sg.Button("閉じる",key="-Auth_exit-"))]]
        elif state.popup_setting: #設定画面
            layout = [[sg.T("設定")],
                      [sg.T("メールアドレス："),sg.T("", key="-email-",expand_x=True)],
                      [sg.T("ユーザー名    ："),sg.T("", key="-UserName-",expand_x=True)],
                      [sg.Button("GitHub",key="-GitHub-"), sg.Button("利用方法",key="-Help-"),
                       sg.Button("アップデート情報",key="-Update-"), sg.Button("お問い合わせ",key="-Contact-"),
                       sg.Button("バグ報告",key="-BugReport-")],
                      # ログインしている場合はログアウトボタンを表示
                      [((sg.Button("ログアウト", key="-Logout-"), 
                        sg.Button("アカウント削除", button_color=("yellow", "red"),key="-Delete-"))
                        if state.isLogin else sg.Button("ログイン", key="-Login-")),
                       ],
                      [sg.Button("閉じる",key="-setting_exit-"),sg.T("バージョン："),sg.T("1.1.0")]]
        elif state.popup_subject_delete:
            layout = [[sg.T("科目修正・削除")],
                      [sg.Table(values=data, headings=header, display_row_numbers=True,
                                      auto_size_columns=True, num_rows=min(25, len(data)),
                                      expand_x=True, expand_y=True)],
                      [sg.T("修正・削除したい科目名を入力してください。")],
                      [sg.T("　　　 科目名："),sg.I("", key="-subject-",expand_x=True)],
                      [sg.T("　　　 単位数："),sg.I("", key="-units_num-",expand_x=True)],
                      [sg.T("評価ポイント："),sg.I("", key="-HPT-",expand_x=True)],
                      [sg.Checkbox("合否科目",default=False, key="-Pass/Fail-")],
                      [sg.Button("修正", key="-Subject_update-"), sg.Button("削除", key="-Subject_delete-"), sg.Button("閉じる",key="-Close-")]]
        else: #GPA計算画面
            layout = [[sg.T("　　　　  累積GPA："),sg.T("0.00", key="-GPA-")],
                      [sg.T("　　　 累積単位数："),sg.T("000", key="-all_total_units_num-")],
                    [sg.T("正課学習ポイント："),sg.T("0.00", key="-SGPT-")],
                    [sg.T("", key = "-txt-")],
                    [sg.T("　　　 科目名："),sg.I("", key="-subject-",expand_x=True)],
                    [sg.T("　　　 単位数："),sg.I("", key="-units_num-",expand_x=True)],
                    [sg.T("評価ポイント："),sg.I("", key="-HPT-",expand_x=True)],
                    [sg.Checkbox("合否科目",default=False, key="-Pass/Fail-")],
                    [sg.T("過去のsubject_grades_data.csvファイル："),sg.I("", key="-inputFilePath-", expand_x=True),
                    sg.FileBrowse("ファイル選択"),],
                    [sg.Button("Submit", key="-Submit-"), sg.Button("Final", key="-Final-"),
                    sg.Button("CSVファイル", key="-CSV-"), sg.Button("ファイルインポート", key="-File_Import-"),
                    sg.Button("科目修正・削除", key="-Subject_UI-"), sg.Button("GPAリセット", key="-GPA_reset-"),
                    sg.Button("設定", key="-Setting-")]]
        self.win = sg.Window("GPA計算", layout, font = (None, 15),
                             finalize=True, resizable = True)

def reload_gui(state, win, data=None, header=None):
    try:
        print("reload_guiが呼び出されました")
        print(state)
        
        gui = GUI(state, data, header)
        new_win = gui.win
        
        if new_win is None:
            print("警告：新しいウィンドウの作成に失敗しました")
        else:
            win.close()
            win = new_win
    except Exception as e:
        print(f"reload_gui でエラーが発生しました: {e}")
    return win