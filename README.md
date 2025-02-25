# KIT_GPA_calc-金沢工業大学GPA計算ソフト
## ソフト利用前に
この計算ソフトは**pandas及びPySimpleGUI,dotenv,firebase-admin,webbrowser,npm**の外部ライブラリを使用しております。pip installにてこのライブラリ群をインストールしてください。また、PySimpleGUIは個人用のみ無償利用が可能となっております。ソフトを利用するためにサインインが必要ですのでサインインをお忘れなく行ってください。
### PySimpleGUI
PySimpleGUIのサインインURLはこちらです。
> https://pysimplegui.com/pricing
### firebase
firebaseの利用は各人でプロジェクトを作成してください。\n
その後、firebaseの設定ファイルを作成してください。\n
設定ファイルは以下のように.envファイルを作成してください。
> ```
> REACT_APP_FIREBASE_API_KEY=your_api_key
> REACT_APP_FIREBASE_AUTH_DOMAIN=your_auth_domain
> REACT_APP_DATABASE_URL=your_database_url
> REACT_APP_DATABASE_PROJECT_ID=your_project_id
> REACT_APP_FIREBASE_STORAGE_BUCKET=your_storage_bucket
> REACT_APP_FIREBASE_MESSAGING_SENDER_ID=your_messaging_sender_id
> REACT_APP_FIREBASE_APP_ID=your_app_id
> REACT_APP_FIREBASE_MEASUREMENTID=
> ```
その後、firebaseの設定ファイル"firebase_api.json"を作成してください。\n
firebaseの設定ファイルは以下のように作成してください。
> ```
> {
> "type": "service_account",
> "project_id": "your_project_id",
> "private_key_id": "your_private_key_id",
> "private_key": "your_private_key",
> "client_email": "your_client_email",
> "client_id": "your_client_id",
> "auth_uri": "https://accounts.google.com/o/oauth2/auth",
> "token_uri": "https://oauth2.googleapis.com/token",
> "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
> "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/your_client_email",
> "universe_domain": "googleapis.com"
> }
> ```
作成したfirebase_api.jsonファイルを**firebase_setting**フォルダに格納してください。
## 利用方法
### GUIからの入力
GUIの科目名、単位数、評価ポイント、合否科目という個所を入力し下部ボタン**Submit**を押下することにより、内部のリストに成績情報として格納される。
※この際単位数には半角数字を評価ポイントにはS,A,B,C,D,F,合,否のいずれかを入力してください。
### 外部ファイルからのインポート
ファイルを選択からこちらの形式にあったCSVファイルを選択してください。その後下部ボタン**ファイルインポート**を押下することにより、内部のリストに成績情報として格納される。
GUIには**過去のsubject_grades_data.csv**と表記されているが、CSVファイルの名称はどのようなものでも可です。
※CSVファイルの形式を以下に示す。
> | 科目名 | 単位数 | 評価ポイント | 合否科目 |
> |--------|-------|-------------|---------|
> | subject a | 2 | A | False |
> | subject b | 1 | 合 | True |
> | : | : | : | : |
### GPA・正課学習ポイントの計算
GUIからの入力または外部ファイルからのインポートを行った後に下部ボタン**Final**を押下することにより、GPA及び正課学習ポイントの計算を行いGUI上に表示される。
### CSVファイルの保存・入力状況の表示
GUIからの入力または外部ファイルからのインポートを行った後に下部ボタン**CSVファイル**を押下することにより、CSVファイルを保存し、CSVファイルをもとに表形式で表示される。
※この時CSVファイルは確認なしで上書き保存されるため、以前のものを残しておきたい場合は自分自身で変更をかけること。
### 成績情報のリセット
下部ボタン**GPAリセット**を押下し確認ボタンを押下することにより、それまでに入力されている成績情報をすべて削除する。\n
そのため、CSVファイルに外部ファイルとして保存しておく必要がある。
### 利用方法
利用方法は**利用方法**から確認できます。
### アップデート情報
アップデート情報は**アップデート情報**から確認できます。
### お問い合わせ
お問い合わせは**お問い合わせ**からお願いします。
### バグ報告
バグ報告は**バグ報告**からお願いします。
### ログイン・ユーザー作成
ログイン・ユーザー作成は**ログイン**からおこなうことができます。\n
ログインをすることによって今後実装する成績情報のデータベースでの管理が行うことができるようになります。
### ログアウト
ログアウトは**ログアウト**からおこなうことができます。\n
ログアウトしても成績情報のデータベースは残っています。
### ユーザー削除
ユーザー削除は**ユーザー削除**からおこなうことができます。\n
ユーザー削除をすることによって今後実装する成績情報をデータベースから完全に削除されますのでご注意ください。\n
ユーザー削除をすることによってログインしているユーザーのログアウトも行われます。
