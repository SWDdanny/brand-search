import os
import requests
import json
import time
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- 設定區 ---
SPREADSHEET_ID = '你的試算表ID'  # 請替換為實際的 ID
SHEET_NAME = 'Sheet1'           # 請確保工作表名稱正確

def get_gspread_service():
    """建立 Google Sheets API 連線"""
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    secret_data = os.getenv("GCP_SERVICE_ACCOUNT")
    if not secret_data:
        raise ValueError("❌ 找不到環境變數 GCP_SERVICE_ACCOUNT")
    
    service_account_info = json.loads(secret_data)
    creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)

def search_company_info(brand_name):
    """
    使用 Serper.dev 搜尋，並鎖定「台灣公司網」
    """
    url = "https://google.serper.dev/search"
    api_key = os.getenv("SERPER_API_KEY")
    
    # 依照需求搜尋 "品牌名稱 公司"
    query = f"{brand_name} 公司"
    
    payload = json.dumps({
        "q": query,
        "gl": "tw",
        "hl": "zh-tw",
        "num": 10  # 增加搜尋數量以確保能找到台灣公司網的連結
    })
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}

    try:
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        res_json = response.json()
        organic = res_json.get("organic", [])
        
        official_title = ""
        phone = ""

        # 1. 優先找「台灣公司網」的結果
        twincn_result = next((item for item in organic if "twincn.com" in item.get("link", "")), None)
        
        if twincn_result:
            # 從台灣公司網的標題提取公司名 (通常格式為: 公司名稱 / 負責人 / ...)
            raw_title = twincn_result.get("title", "")
            official_title = raw_title.split('｜')[0].split('|')[0].split(' - ')[0].strip()
            
            # 嘗試從摘要中用正規表達式抓取電話
            snippet = twincn_result.get("snippet", "")
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()
        
        # 2. 如果沒找到台灣公司網，退而求其次使用第一筆搜尋結果
        if not official_title and organic:
            official_title = organic[0].get("title", "").split('-')[0].split('|')[0].strip()
            # 從摘要找電話
            snippet = organic[0].get("snippet", "")
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()

        return official_title, phone
    except Exception as e:
        print(f"❌ 搜尋 {brand_name} 失敗: {e}")
        return "", ""

def main():
    service = get_gspread_service()
    sheet = service.spreadsheets()

    # 讀取 A 到 K 欄
    range_to_read = f"{SHEET_NAME}!A2:K"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
    rows = result.get('values', [])

    if not rows:
        print("📭 找不到任何資料。")
        return

    for i, row in enumerate(rows):
        # 補足空欄位
        while len(row) < 11:
            row.append("")

        # 欄位對應：C欄是索引 2, H欄是索引 7, J欄是索引 9, K欄是索引 10
        brand_name = row[2]     # C欄：品牌名稱
        status = row[7]         # H欄：狀態
        existing_title = row[9] # J欄：正式抬頭

        # 篩選：狀態為「已分配」且 J 欄為空
        if status == "已分配" and not existing_title:
            if not brand_name:
                continue
                
            print(f"🔎 正在搜尋品牌: {brand_name}...")
            official_title, phone = search_company_info(brand_name)
            
            if official_title:
                row_num = i + 2
                update_range = f"{SHEET_NAME}!J{row_num}:K{row_num}"
                update_body = {
                    "values": [[official_title, phone]]
                }
                
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="RAW",
                    body=update_body
                ).execute()
                
                print(f"✅ 已填入: {official_title} | 電話: {phone}")
            
            time.sleep(1) # 稍微冷卻避免觸發頻率限制

    print("🏁 處理完畢！")

if __name__ == "__main__":
    main()
