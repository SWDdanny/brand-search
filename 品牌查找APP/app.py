import streamlit as st
import datetime
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path
import pickle

# 匯入我們剛建立的常數檔案 [cite: 1, 6]
from constants import CALENDAR_DATA, MAJOR_EVENTS

# --- 設定區 ---
SPREADSHEET_ID = '1CIjHi8dImHdLmNdzSMXh0qf1pE1KZHFBszS4SDdqVOg'
CLIENT_SECRET_FILE = 'client_secret_2_740099921822-cl516v2gti8tkev687hftqcas93471km.apps.googleusercontent.com.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

st.set_page_config(page_title="品牌開發爬蟲系統", layout="centered")
st.title("🔍 品牌開發自動化搜尋")

# 1. 自動計算下個月 
today = datetime.date.today()
next_month_val = (today.month % 12) + 1
month_key = f"{next_month_val}月"

st.subheader(f"📅 當前目標：搜尋 {month_key} 的潛在客戶")

# 2. 從 constants 讀取對應資料 
target_data = CALENDAR_DATA.get(month_key, {"events": [], "industries": []})
events = target_data["events"]
industries = target_data["industries"]

# 3. 網頁介面選擇
col1, col2 = st.columns(2)
with col1:
    selected_event = st.selectbox("選擇當月檔期", events)
with col2:
    selected_industry = st.selectbox("選擇目標產業", industries)

# 4. 排除設定
exclude_sites = "-site:104.com.tw -site:1111.com.tw -site:shopee.tw -site:momo.com.tw"
search_query = f"2026 {month_key} {selected_event} {selected_industry} 品牌 電話 {exclude_sites}"

st.info(f"🚀 預計搜尋關鍵字：\n`{search_query}`")

# 5. 寫入 Google Sheets 的功能
def write_to_sheets(results):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    body = {'values': results}
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="A1",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()

# 6. 執行按鈕
if st.button("開始抓取並存入 Google Sheets"):
    # 在這裡模擬抓取到的資料 (實際 Selenium 建議在本機跑)
    mock_results = [
        [str(datetime.date.today()), selected_event, selected_industry, "範例品牌名稱", "02-23456789", "搜尋結果摘要..."]
    ]
    
    try:
        write_to_sheets(mock_results)
        st.success(f"成功！資料已寫入試算表 ID: {SPREADSHEET_ID}")
        st.balloons()
    except Exception as e:
        st.error(f"寫入失敗，錯誤訊息：{e}")
