import time
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import pytz

# Googleスプレッドシートの設定
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Streamlit secretsからcredentialsとSPREADSHEET_KEYを取得
credentials_info = st.secrets["gcp_service_account"]
spreadsheet_key = st.secrets["spreadsheet_key"]["spreadsheet_key"]
credentials = Credentials.from_service_account_info(credentials_info, scopes=scope)
gc = gspread.authorize(credentials)
workbook = gc.open_by_key(spreadsheet_key)
sheet = workbook.sheet1
question_sheet = workbook.worksheet("質問")  # 質問シートを追加

# セッションステートの初期化
if 'confirm_overwrite' not in st.session_state:
    st.session_state.confirm_overwrite = False
    st.session_state.row_number = None
    st.session_state.col_number = None
    st.session_state.answer = None
    st.session_state.existing_answer = None

def update_sheet(username, question_num, answer):
    # スプレッドシートにデータを書き込む関数
    cell = sheet.find(username)
    if cell:
        # ユーザー名が既に存在する場合、その行を更新
        row_number = cell.row
        existing_answer = sheet.cell(row_number, question_num + 1).value
        if existing_answer:
            st.session_state.confirm_overwrite = True
            st.session_state.row_number = row_number
            st.session_state.col_number = question_num + 1
            st.session_state.answer = answer
            st.session_state.existing_answer = existing_answer
        else:
            sheet.update_cell(row_number, question_num + 1, answer)
            st.success("データをスプレッドシートに送信しました。")
    else:
        # 新しいユーザーの場合、新しい行を追加
        new_row = [username] + [''] * (sheet.col_count - 1)
        sheet.append_row(new_row)
        row_number = sheet.find(username).row
        sheet.update_cell(row_number, question_num + 1, answer)
        st.success("データをスプレッドシートに送信しました。")

def add_question(name, question):
    # 現在の日時を東京タイムゾーンで取得
    tz = pytz.timezone('Asia/Tokyo')
    current_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')
    # 質問シートにデータを書き込む関数
    question_sheet.append_row([name, question, current_time])
    st.success("質問をスプレッドシートに送信しました。")

# プルダウンメニューの選択肢を作成
options = [
    "ワーク2-1 プロンプト",
    "ワーク2-1 ChatGPTの回答",
    "ワーク2-2 プロンプト",
    "ワーク2-2 ChatGPTの回答",
    "ワーク2-3 プロンプト",
    "ワーク2-3 ChatGPTの回答",
    "ワーク2-4 プロンプト",
    "ワーク2-4 ChatGPTの回答",
    "ワーク2-5 プロンプト",
    "ワーク2-5 ChatGPTの回答",
    "ワーク2-6 プロンプト",
    "ワーク2-6 ChatGPTの回答",
    "ワーク3-1 プロンプト",
    "ワーク3-1 ChatGPTの回答",
    "ワーク4-1 プロンプト",
    "ワーク4-1 ChatGPTの回答",
    "ワーク4-2 プロンプト",
    "ワーク4-2 ChatGPTの回答"
]

# Streamlit UI
with st.form("user_input"):
    username = st.text_input("ユーザー名")
    question_num = st.selectbox("質問番号", options)
    answer = st.text_area("回答", height=300)
    submitted = st.form_submit_button("送信")

if submitted:
    if not username:
        st.error("ユーザー名を入力してください。")
    else:
        # 選択された質問番号から実際のインデックスを取得
        question_index = options.index(question_num) + 1
        update_sheet(username, question_index, answer)

if st.session_state.confirm_overwrite:
    st.warning("既存の回答があります。上書きしてもよろしいですか？")
    st.write(f"既存の回答: {st.session_state.existing_answer}")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("はい", key="yes"):
            sheet.update_cell(st.session_state.row_number, st.session_state.col_number, st.session_state.answer)
            st.success("データをスプレッドシートに上書きしました。")
            time.sleep(2)
            st.session_state.confirm_overwrite = False
            st.rerun()  # ページをリロードしてUIを更新
    with col2:
        if st.button("いいえ", key="no"):
            st.info("操作がキャンセルされました。")
            time.sleep(1)
            st.session_state.confirm_overwrite = False
            st.rerun()  # ページをリロードしてUIを更新

# 質問フォームのUI
st.header("質問フォーム")
with st.form("question_form"):
    name = st.text_input("名前")
    question = st.text_area("質問", height=100)
    question_submitted = st.form_submit_button("質問を送信")

if question_submitted:
    if not name or not question:
        st.error("名前と質問を入力してください。")
    else:
        add_question(name, question)
