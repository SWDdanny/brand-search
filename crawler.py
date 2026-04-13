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
    # 移除常見干擾標記與網址描述
    name = raw_title.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    # 移除括號內的文字 (例如: 統一編號)
    name = re.sub(r'\(.*\)', '', name).strip()
    return name

def is_valid_company_name(name):
    """檢查字串是否包含公司關鍵字，並排除描述性語句"""
    keywords = ["公司", "集團", "行號", "有限", "工作室", "企業"]
    invalid_keywords = ["一家", "身處", "領域", "科技公司", "官網", "介紹"]
    
    # 必須包含正確關鍵字
    has_keyword = any(k in name for k in keywords)
    # 不得包含描述性虛詞
    not_descriptive = not any(ik in name for ik in invalid_keywords)
    # 正式名稱通常不會太長
    not_too_long = len(name) < 25
    
    return has_keyword and not_descriptive and not_too_long

def search_company_info(brand_name):
    print(f"🔎 步驟 1: 查找品牌正式抬頭 -> {brand_name}")
    
    # 策略 A: 強制優先從「台灣公司網」搜尋
    results_step1 = serper_request(f"{brand_name} site:twincn.com")
    
    official_title = ""
    phone = "查無資料"

    if results_step1:
        for item in results_step1:
            temp_name = clean_company_name(item.get("title", ""))
            if is_valid_company_name(temp_name):
                official_title = temp_name
                print(f"🎯 從台灣公司網命中: {official_title}")
                break
    
    # 策略 B: 如果台灣公司網沒結果，進行廣泛搜尋但加強條件
    if not official_title:
        print(f"⚠️ 台灣公司網未命中，嘗試廣泛搜尋正確抬頭...")
        results_wide = serper_request(f"{brand_name} 台灣正式登記公司名稱")
        for item in results_wide:
            temp_name = clean_company_name(item.get("title", ""))
            if is_valid_company_name(temp_name):
                official_title = temp_name
                break

    # 若最終還是沒找到合適的抬頭
    if not official_title:
        return "查無品牌", "查無資料"

    # 步驟 2: 查找電話
    print(f"🔎 步驟 2: 查找電話 -> {official_title}")
    # 針對確定的正式抬頭精確搜尋
    results_step2 = serper_request(f"{official_title} site:twincn.com")
    
    found_phone = False
    for item in results_step2:
        snippet = item.get("snippet", "")
        phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
        if phone_match:
            phone = phone_match.group()
            found_phone = True
            break
    
    if not found_phone:
        # 廣泛搜尋電話
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

        # 判定邏輯：狀態已分配 且 J欄位空，且不曾被標記為查無資料
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
            time.sleep(1.5) # 稍微加長間隔提高穩定度

    print("🏁 任務結束。")

if __name__ == "__main__":
    main()
