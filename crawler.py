import os
import requests
import json
import time
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- 設定區 ---
SPREADSHEET_ID = '1jb7MZ5w00zNs3T_I7lxT24nEChudAUnUnpXLm77sOXU' 
SHEET_NAME = '品牌名單'          

def get_gspread_service():
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    secret_data = os.getenv("GCP_SERVICE_ACCOUNT")
    if not secret_data:
        raise ValueError("❌ 找不到環境變數 GCP_SERVICE_ACCOUNT")
    
    service_account_info = json.loads(secret_data)
    creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return build('sheets', 'v4', credentials=creds)

def search_company_info(brand_name):
    url = "https://google.serper.dev/search"
    api_key = os.getenv("SERPER_API_KEY")
    
    # 搜尋 "品牌名稱 公司"
    query = f"{brand_name} 公司"
    
    payload = json.dumps({
        "q": query,
        "gl": "tw",
        "hl": "zh-tw",
        "num": 10 
    })
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}

    try:
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        res_json = response.json()
        organic = res_json.get("organic", [])
        
        official_title = ""
        phone = ""

        # 優先找「台灣公司網」
        twincn_result = next((item for item in organic if "twincn.com" in item.get("link", "")), None)
        
        if twincn_result:
            raw_title = twincn_result.get("title", "")
            official_title = raw_title.split('｜')[0].split('|')[0].split(' - ')[0].strip()
            snippet = twincn_result.get("snippet", "")
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()
        
        # 若沒找到，用第一筆
        if not official_title and organic:
            official_title = organic[0].get("title", "").split('-')[0].split('|')[0].strip()
            snippet = organic[0].get("snippet", "")
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()

        # 如果最終還是空的，回傳標記
        return (official_title or "查無資料"), (phone or "查無資料")
    except Exception as e:
        print(f"❌ 搜尋 {brand_name} 出錯: {e}")
        return "查無資料", "查無資料"

def main():
    service = get_gspread_service()
    sheet = service.spreadsheets()

    range_to_read = f"{SHEET_NAME}!A2:K"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
    rows = result.get('values', [])

    if not rows:
        print("📭 找不到資料。")
        return

    for i, row in enumerate(rows):
        while len(row) < 11:
            row.append("")

        brand_name = row[2]      # C欄
        status = row[7].strip()  # H欄
        existing_title = row[9].strip() # J欄

        # 修改後的判定邏輯：
        # 狀態是「已分配」 且 (J欄是空的 OR J欄內容是 "查無資料")
        if status == "已分配" and (not existing_title or existing_title == "查無資料"):
            if not brand_name or brand_name == "查無資料":
                continue
                
            print(f"🔎 正在重新查找: {brand_name}...")
            official_title, phone = search_company_info(brand_name)
            
            row_num = i + 2
            update_range = f"{SHEET_NAME}!J{row_num}:K{row_num}"
            update_body = {"values": [[official_title, phone]]}
            
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=update_range,
                valueInputOption="RAW",
                body=update_body
            ).execute()
            
            print(f"✅ 結果: {official_title} | {phone}")
            time.sleep(1) 

    print("🏁 處理完畢！")

if __name__ == "__main__":
    main()
