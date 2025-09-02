import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os

class GUI:
    def __init__(self, state, data=None, header=None):
        self.state = state
        self.data = data
        self.header = header
        self.win = None
        self.widgets = {}
        self.create_window()

    def create_window(self):
        if self.state.popup_login:  # ログイン画面
            self.create_login_window()
        elif self.state.popup_setting:  # 設定画面
            self.create_setting_window()
        elif self.state.popup_subject_delete:  # 科目修正・削除画面
            self.create_subject_edit_window()
        else:  # GPA計算画面
            self.create_main_window()

    def create_main_window(self):
        self.win = tk.Tk()
        self.win.title("GPA計算")
        self.win.geometry("800x600")
        
        # メインフレーム
        main_frame = ttk.Frame(self.win, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # GPA表示部分
        gpa_frame = ttk.LabelFrame(main_frame, text="GPA情報", padding="5")
        gpa_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(gpa_frame, text="累積GPA:").grid(row=0, column=0, sticky=tk.W)
        self.widgets["-GPA-"] = ttk.Label(gpa_frame, text="0.00")
        self.widgets["-GPA-"].grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        
        ttk.Label(gpa_frame, text="卒業単位数:").grid(row=0, column=2, sticky=tk.W)
        self.widgets["-Graduation_total_units_num-"] = ttk.Label(gpa_frame, text="0単位")
        self.widgets["-Graduation_total_units_num-"].grid(row=0, column=3, sticky=tk.W, padx=(5, 20))
        
        ttk.Label(gpa_frame, text="累積単位数:").grid(row=1, column=0, sticky=tk.W)
        self.widgets["-all_total_units_num-"] = ttk.Label(gpa_frame, text="0単位")
        self.widgets["-all_total_units_num-"].grid(row=1, column=1, sticky=tk.W, padx=(5, 20))
        
        ttk.Label(gpa_frame, text="正課学習ポイント:").grid(row=1, column=2, sticky=tk.W)
        self.widgets["-SGPT-"] = ttk.Label(gpa_frame, text="0.00")
        self.widgets["-SGPT-"].grid(row=1, column=3, sticky=tk.W, padx=(5, 20))
        
        # メッセージ表示
        self.widgets["-txt-"] = ttk.Label(main_frame, text="", foreground="red")
        self.widgets["-txt-"].grid(row=1, column=0, columnspan=2, pady=5)
        
        # 入力部分
        input_frame = ttk.LabelFrame(main_frame, text="科目入力", padding="5")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(input_frame, text="科目名:").grid(row=0, column=0, sticky=tk.W)
        self.widgets["-subject-"] = ttk.Entry(input_frame, width=30)
        self.widgets["-subject-"].grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0))
        
        ttk.Label(input_frame, text="単位数:").grid(row=1, column=0, sticky=tk.W)
        self.widgets["-units_num-"] = ttk.Entry(input_frame, width=30)
        self.widgets["-units_num-"].grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0))
        
        ttk.Label(input_frame, text="評価ポイント:").grid(row=2, column=0, sticky=tk.W)
        self.widgets["-HPT-"] = ttk.Entry(input_frame, width=30)
        self.widgets["-HPT-"].grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 0))
        
        # チェックボックス
        checkbox_frame = ttk.Frame(input_frame)
        checkbox_frame.grid(row=3, column=0, columnspan=2, pady=5)
        
        self.widgets["-Pass/Fail-"] = tk.BooleanVar()
        ttk.Checkbutton(checkbox_frame, text="合否科目", variable=self.widgets["-Pass/Fail-"]).grid(row=0, column=0, sticky=tk.W)
        
        self.widgets["-Teacher-"] = tk.BooleanVar()
        ttk.Checkbutton(checkbox_frame, text="教職科目", variable=self.widgets["-Teacher-"]).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # ファイル選択部分
        file_frame = ttk.LabelFrame(main_frame, text="ファイルインポート", padding="5")
        file_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(file_frame, text="CSVファイル:").grid(row=0, column=0, sticky=tk.W)
        self.widgets["-inputFilePath-"] = ttk.Entry(file_frame, width=40)
        self.widgets["-inputFilePath-"].grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5))
        
        ttk.Button(file_frame, text="ファイル選択", command=self.browse_file).grid(row=0, column=2, sticky=tk.W)
        
        # ボタン部分
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Submit", command=self.submit_clicked).grid(row=0, column=0, padx=2)
        ttk.Button(button_frame, text="Final", command=self.final_clicked).grid(row=0, column=1, padx=2)
        ttk.Button(button_frame, text="CSVファイル", command=self.csv_clicked).grid(row=0, column=2, padx=2)
        ttk.Button(button_frame, text="ファイルインポート", command=self.file_import_clicked).grid(row=0, column=3, padx=2)
        ttk.Button(button_frame, text="科目修正・削除", command=self.subject_ui_clicked).grid(row=0, column=4, padx=2)
        ttk.Button(button_frame, text="GPAリセット", command=self.gpa_reset_clicked).grid(row=0, column=5, padx=2)
        ttk.Button(button_frame, text="設定", command=self.setting_clicked).grid(row=0, column=6, padx=2)
        
        # グリッドの重み設定
        self.win.columnconfigure(0, weight=1)
        self.win.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(1, weight=1)
        file_frame.columnconfigure(1, weight=1)

    def create_login_window(self):
        self.win = tk.Toplevel()
        self.win.title("新規登録" if self.state.popup_signup else "ログイン")
        self.win.geometry("400x300")
        self.win.grab_set()  # モーダルウィンドウにする
        
        main_frame = ttk.Frame(self.win, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="新規登録" if self.state.popup_signup else "ログイン", font=("", 16, "bold")).pack(pady=10)
        
        self.widgets["-txt-"] = ttk.Label(main_frame, text="", foreground="red")
        self.widgets["-txt-"].pack(pady=5)
        
        ttk.Label(main_frame, text="メールアドレス:").pack(anchor=tk.W)
        self.widgets["-email-"] = ttk.Entry(main_frame, width=40)
        self.widgets["-email-"].pack(pady=5, fill=tk.X)
        
        if self.state.popup_signup:
            ttk.Label(main_frame, text="ユーザー名:").pack(anchor=tk.W)
            self.widgets["-UserName-"] = ttk.Entry(main_frame, width=40)
            self.widgets["-UserName-"].pack(pady=5, fill=tk.X)
        
        ttk.Label(main_frame, text="パスワード:").pack(anchor=tk.W)
        self.widgets["-password-"] = ttk.Entry(main_frame, width=40, show="*")
        self.widgets["-password-"].pack(pady=5, fill=tk.X)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        if self.state.popup_signup:
            ttk.Button(button_frame, text="登録", command=self.auth_create_clicked).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="戻る", command=self.back_auth_clicked).pack(side=tk.LEFT, padx=5)
        else:
            ttk.Button(button_frame, text="ログイン", command=self.auth_clicked).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="新規登録", command=self.user_create_clicked).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="閉じる", command=self.auth_exit_clicked).pack(side=tk.LEFT, padx=5)

    def create_setting_window(self):
        self.win = tk.Toplevel()
        self.win.title("設定")
        self.win.geometry("500x400")
        self.win.grab_set()  # モーダルウィンドウにする
        
        main_frame = ttk.Frame(self.win, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="設定", font=("", 16, "bold")).pack(pady=10)
        
        # ユーザー情報表示
        info_frame = ttk.LabelFrame(main_frame, text="ユーザー情報", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(info_frame, text="メールアドレス:").grid(row=0, column=0, sticky=tk.W)
        self.widgets["-email-"] = ttk.Label(info_frame, text=self.state.user_email)
        self.widgets["-email-"].grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        ttk.Label(info_frame, text="ユーザー名:").grid(row=1, column=0, sticky=tk.W)
        self.widgets["-UserName-"] = ttk.Label(info_frame, text=self.state.user_name)
        self.widgets["-UserName-"].grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # リンクボタン
        link_frame = ttk.LabelFrame(main_frame, text="リンク", padding="10")
        link_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(link_frame, text="GitHub", command=self.github_clicked).grid(row=0, column=0, padx=5)
        ttk.Button(link_frame, text="利用方法", command=self.help_clicked).grid(row=0, column=1, padx=5)
        ttk.Button(link_frame, text="アップデート情報", command=self.update_clicked).grid(row=0, column=2, padx=5)
        ttk.Button(link_frame, text="お問い合わせ", command=self.contact_clicked).grid(row=1, column=0, padx=5)
        ttk.Button(link_frame, text="バグ報告", command=self.bug_report_clicked).grid(row=1, column=1, padx=5)
        
        # 認証ボタン
        auth_frame = ttk.Frame(main_frame)
        auth_frame.pack(pady=10)
        
        if self.state.isLogin:
            ttk.Button(auth_frame, text="ログアウト", command=self.logout_clicked).pack(side=tk.LEFT, padx=5)
            ttk.Button(auth_frame, text="アカウント削除", command=self.delete_clicked).pack(side=tk.LEFT, padx=5)
        else:
            ttk.Button(auth_frame, text="ログイン", command=self.login_clicked).pack(side=tk.LEFT, padx=5)
        
        # 閉じるボタンとバージョン
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        ttk.Button(bottom_frame, text="閉じる", command=self.setting_exit_clicked).pack(side=tk.LEFT)
        ttk.Label(bottom_frame, text="バージョン：1.1.0").pack(side=tk.RIGHT)

    def create_subject_edit_window(self):
        self.win = tk.Toplevel()
        self.win.title("科目修正・削除")
        self.win.geometry("800x600")
        self.win.grab_set()  # モーダルウィンドウにする
        
        main_frame = ttk.Frame(self.win, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="科目修正・削除", font=("", 16, "bold")).pack(pady=10)
        
        # テーブル表示
        table_frame = ttk.LabelFrame(main_frame, text="科目一覧", padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Treeviewでテーブルを作成
        columns = self.header if self.header else ["科目名", "単位数", "評価ポイント", "合否科目", "教職科目"]
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # データを挿入
        if self.data:
            for row in self.data:
                self.tree.insert('', tk.END, values=row)
        
        # 入力部分
        input_frame = ttk.LabelFrame(main_frame, text="修正・削除", padding="5")
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(input_frame, text="修正・削除したい科目名を入力してください。").pack(anchor=tk.W)
        
        # 入力フィールド用のフレーム
        fields_frame = ttk.Frame(input_frame)
        fields_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(fields_frame, text="科目名:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.widgets["-subject-"] = ttk.Entry(fields_frame, width=30)
        self.widgets["-subject-"].grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        
        ttk.Label(fields_frame, text="単位数:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.widgets["-units_num-"] = ttk.Entry(fields_frame, width=10)
        self.widgets["-units_num-"].grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(0, 20))
        
        ttk.Label(fields_frame, text="評価ポイント:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        self.widgets["-HPT-"] = ttk.Entry(fields_frame, width=10)
        self.widgets["-HPT-"].grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        
        # チェックボックス
        checkbox_frame = ttk.Frame(fields_frame)
        checkbox_frame.grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=(20, 0))
        
        self.widgets["-Pass/Fail-"] = tk.BooleanVar()
        ttk.Checkbutton(checkbox_frame, text="合否科目", variable=self.widgets["-Pass/Fail-"]).pack(side=tk.LEFT)
        
        self.widgets["-Teacher-"] = tk.BooleanVar()
        ttk.Checkbutton(checkbox_frame, text="教職科目", variable=self.widgets["-Teacher-"]).pack(side=tk.LEFT, padx=(20, 0))
        
        # ボタン
        button_frame = ttk.Frame(input_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="修正", command=self.subject_update_clicked).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="削除", command=self.subject_delete_clicked).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="閉じる", command=self.close_clicked).pack(side=tk.LEFT, padx=5)
        
        fields_frame.columnconfigure(1, weight=1)
        fields_frame.columnconfigure(3, weight=1)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.widgets["-inputFilePath-"].delete(0, tk.END)
            self.widgets["-inputFilePath-"].insert(0, filename)

    # イベントハンドラーメソッド
    def submit_clicked(self):
        self.win.event_generate("<<Submit>>")

    def final_clicked(self):
        self.win.event_generate("<<Final>>")

    def csv_clicked(self):
        self.win.event_generate("<<CSV>>")

    def file_import_clicked(self):
        self.win.event_generate("<<File_Import>>")

    def subject_ui_clicked(self):
        self.win.event_generate("<<Subject_UI>>")

    def gpa_reset_clicked(self):
        self.win.event_generate("<<GPA_reset>>")

    def setting_clicked(self):
        self.win.event_generate("<<Setting>>")

    def auth_clicked(self):
        self.win.event_generate("<<Auth>>")

    def auth_create_clicked(self):
        self.win.event_generate("<<Auth_Create>>")

    def back_auth_clicked(self):
        self.win.event_generate("<<back_Auth>>")

    def auth_exit_clicked(self):
        self.win.event_generate("<<Auth_exit>>")

    def user_create_clicked(self):
        self.win.event_generate("<<UserCreate>>")

    def setting_exit_clicked(self):
        self.win.event_generate("<<setting_exit>>")

    def github_clicked(self):
        self.win.event_generate("<<GitHub>>")

    def help_clicked(self):
        self.win.event_generate("<<Help>>")

    def update_clicked(self):
        self.win.event_generate("<<Update>>")

    def contact_clicked(self):
        self.win.event_generate("<<Contact>>")

    def bug_report_clicked(self):
        self.win.event_generate("<<BugReport>>")

    def login_clicked(self):
        self.win.event_generate("<<Login>>")

    def logout_clicked(self):
        self.win.event_generate("<<Logout>>")

    def delete_clicked(self):
        self.win.event_generate("<<Delete>>")

    def subject_update_clicked(self):
        self.win.event_generate("<<Subject_update>>")

    def subject_delete_clicked(self):
        self.win.event_generate("<<Subject_delete>>")

    def close_clicked(self):
        self.win.event_generate("<<Close>>")

    def get_values(self):
        """PySimpleGUIのwin.read()の代わりに使用するメソッド"""
        values = {}
        for key, widget in self.widgets.items():
            if isinstance(widget, ttk.Entry):
                values[key] = widget.get()
            elif isinstance(widget, tk.BooleanVar):
                values[key] = widget.get()
            elif isinstance(widget, ttk.Label):
                values[key] = widget.cget("text")
        return values

    def update_widget(self, key, value):
        """PySimpleGUIのwin[key].update()の代わりに使用するメソッド"""
        if key in self.widgets:
            widget = self.widgets[key]
            if isinstance(widget, ttk.Entry):
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
            elif isinstance(widget, tk.BooleanVar):
                widget.set(bool(value))
            elif isinstance(widget, ttk.Label):
                widget.config(text=str(value))

    def close(self):
        if self.win:
            self.win.destroy()

def reload_gui(state, win, data=None, header=None):
    try:
        print("reload_guiが呼び出されました")
        print(state)
        
        # 既存のウィンドウを閉じる
        if win:
            if hasattr(win, 'destroy'):
                win.destroy()
            elif hasattr(win, 'close'):
                win.close()
        
        # 新しいGUIを作成
        gui = GUI(state, data, header)
        new_win = gui.win
        
        if new_win is None:
            print("警告：新しいウィンドウの作成に失敗しました")
        else:
            win = new_win
    except Exception as e:
        print(f"reload_gui でエラーが発生しました: {e}")
    return win