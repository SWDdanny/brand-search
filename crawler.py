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

def serper_request(query):
    url = "https://google.serper.dev/search"
    api_key = os.getenv("SERPER_API_KEY")
    payload = json.dumps({"q": query, "gl": "tw", "hl": "zh-tw", "num": 10})
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    try:
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        return response.json().get("organic", [])
    except Exception as e:
        print(f"❌ API 請求失敗: {e}")
        return []

def clean_company_name(raw_title):
    # 移除常見干擾標記
    name = raw_title.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    return name

def is_valid_company_name(name):
    """檢查字串是否包含公司關鍵字"""
    keywords = ["公司", "集團", "行號", "有限", "社企", "工作室", "企業"]
    return any(k in name for k in keywords)

def search_company_info(brand_name):
    print(f"🔎 步驟 1: 查找品牌正式抬頭 -> {brand_name}")
    results_step1 = serper_request(f"{brand_name} 台灣公司")
    
    official_title = ""
    phone = "查無資料"

    if results_step1:
        # 優先找「台灣公司網」的結果
        tw_inc_res = next((item for item in results_step1 if "twincn.com" in item.get("link", "")), None)
        
        if tw_inc_res:
            temp_name = clean_company_name(tw_inc_res.get("title", ""))
            if is_valid_company_name(temp_name):
                official_title = temp_name
        
        # 如果沒找到台灣公司網，找其他標題帶有公司關鍵字的結果
        if not official_title:
            for item in results_step1:
                temp_name = clean_company_name(item.get("title", ""))
                if is_valid_company_name(temp_name):
                    official_title = temp_name
                    break

    # 若最終還是沒找到合適的抬頭
    if not official_title:
        return "查無品牌", "查無資料"

    # 步驟 2: 查找電話
    print(f"🔎 步驟 2: 查找電話 -> {official_title}")
    results_step2 = serper_request(f"{official_title} site:twincn.com")
    
    # 嘗試從 snippet 中抓取電話
    found_phone = False
    for item in results_step2:
        snippet = item.get("snippet", "")
        phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
        if phone_match:
            phone = phone_match.group()
            found_phone = True
            break
    
    # 如果台灣公司網沒電話，搜尋該公司名稱廣泛查找
    if not found_phone:
        results_step3 = serper_request(f"{official_title} 電話")
        for item in results_step3[:3]:
            snippet = item.get("snippet", "")
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()
                break

    return official_title, phone

def main():
    service = get_gspread_service()
    sheet = service.spreadsheets()

    range_to_read = f"{SHEET_NAME}!A2:K"
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
        rows = result.get('values', [])
    except Exception as e:
        print(f"❌ 讀取失敗: {e}")
        return

    if not rows:
        print("📭 找不到資料。")
        return

    for i, row in enumerate(rows):
        while len(row) < 11:
            row.append("")

        brand_name = row[2].strip()      # C欄
        status = row[7].strip()          # H欄
        existing_title = row[9].strip()  # J欄
        existing_phone = row[10].strip() # K欄

        # 判定邏輯：狀態已分配 且 (J欄與K欄只要有一個是空的，且都沒有查無資訊的標記)
        # 如果 J 是 '查無品牌' 或 K 是 '查無資料'，就不再搜尋
        is_processed = any(x in [existing_title, existing_phone] for x in ["查無品牌", "查無資料"])
        
        if status == "已分配" and not existing_title and not is_processed:
            if not brand_name:
                continue
                
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
            
            print(f"✅ 回填成功: {brand_name} -> {official_title} | {phone}")
            time.sleep(1.2)

    print("🏁 任務結束。")

if __name__ == "__main__":
    main()
