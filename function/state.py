# グローバル状態管理用の変数
isLogin = False
popup_login = False
popup_signup = False
popup_setting = False
popup_csv = False
popup_subject_delete = False
user_name = ""
user_email = ""
user_id = ""
idToken = ""
current_user_email = ""
current_user_name = ""

def update_state(login=None, p_login=None, p_signup=None, p_setting=None, p_csv=None, p_subject_delete=None, user_name_param=None, user_email_param=None, user_id_param=None, idToken_param=None):
    global isLogin, popup_login, popup_signup, popup_setting, popup_csv, popup_subject_delete, user_name, user_email, user_id, idToken, current_user_email, current_user_name

    print(f"更新前: isLogin={isLogin}, popup_login={popup_login}, popup_signup={popup_signup}, popup_setting={popup_setting}, popup_csv={popup_csv}, popup_subject_delete={popup_subject_delete}, user_name={user_name}, user_email={user_email}, user_id={user_id}")
    if login is not None:
        isLogin = login
    if p_login is not None:
        popup_login = p_login
    if p_signup is not None:
        popup_signup = p_signup
    if p_setting is not None:
        popup_setting = p_setting
    if p_csv is not None:
        popup_csv = p_csv
    if p_subject_delete is not None:
        popup_subject_delete = p_subject_delete
    if user_name_param is not None:
        user_name = user_name_param
        current_user_name = user_name_param  # パラメータも保存
    if user_email_param is not None:
        user_email = user_email_param
        current_user_email = user_email_param  # パラメータも保存
    if user_id_param is not None:
        user_id = user_id_param
    if idToken_param is not None:
        idToken = idToken_param

    print(f"更新後: isLogin={isLogin}, popup_login={popup_login}, popup_signup={popup_signup}, popup_setting={popup_setting}, popup_csv={popup_csv}, popup_subject_delete={popup_subject_delete}, user_name={user_name}, user_email={user_email}, user_id={user_id}")