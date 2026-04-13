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
    """封裝 Serper API 請求"""
    url = "https://google.serper.dev/search"
    api_key = os.getenv("SERPER_API_KEY")
    payload = json.dumps({"q": query, "gl": "tw", "hl": "zh-tw", "num": 8})
    headers = {'X-API-KEY': api_key, 'Content-Type': 'application/json'}
    try:
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        return response.json().get("organic", [])
    except Exception as e:
        print(f"❌ API 請求失敗: {e}")
        return []

def clean_company_name(raw_title):
    """清理標題字串，提取可能的公司正式名稱"""
    # 移除常見的干擾字眼
    name = raw_title.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    # 只要包含「公司」、「有限」、「行號」等關鍵字，通常較準確
    return name

def search_company_info(brand_name):
    print(f"🔎 步驟 1: 查找公司正確抬頭 -> {brand_name}")
    # 改用 "品牌名稱 台灣公司" 增加命中率
    results_step1 = serper_request(f"{brand_name} 台灣公司")
    
    official_title = ""
    phone = "查無資料"

    if results_step1:
        # 優先找標題中帶有「股份有限公司」或「有限公司」的結果
        for item in results_step1:
            title = item.get("title", "")
            if "公司" in title or "有限公司" in title:
                official_title = clean_company_name(title)
                break
        
        # 如果都沒找到帶「公司」字眼的，取第一筆
        if not official_title:
            official_title = clean_company_name(results_step1[0].get("title", ""))

    # 步驟 2: 如果有找到抬頭，專攻台灣公司網找電話
    if official_title and official_title != "查無資料":
        print(f"🔎 步驟 2: 進入台灣公司網查找電話 -> {official_title}")
        results_step2 = serper_request(f"{official_title} site:twincn.com")
        
        # 尋找 twincn 的結果
        twincn_item = next((i for i in results_step2 if "twincn.com" in i.get("link", "")), None)
        
        if twincn_item:
            snippet = twincn_item.get("snippet", "")
            # 正規表達式找電話：例如 02-12345678 或 04-1234567
            phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
            if phone_match:
                phone = phone_match.group()
        
        # 如果 twincn 沒電話，嘗試從一般搜尋結果找
        if phone == "查無資料":
            for item in results_step2[:3]: # 只看前三名
                snippet = item.get("snippet", "")
                phone_match = re.search(r'0\d{1,2}-\d{6,8}', snippet)
                if phone_match:
                    phone = phone_match.group()
                    break

    return (official_title or "查無資料"), phone

def main():
    service = get_gspread_service()
    sheet = service.spreadsheets()

    range_to_read = f"{SHEET_NAME}!A2:K"
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
        rows = result.get('values', [])
    except Exception as e:
        print(f"❌ 讀取試算表失敗，請檢查 ID 或權限: {e}")
        return

    if not rows:
        print("📭 找不到資料。")
        return

    for i, row in enumerate(rows):
        while len(row) < 11:
            row.append("")

        brand_name = row[2]      # C欄
        status = row[7].strip()  # H欄
        existing_title = row[9].strip() # J欄

        if status == "已分配" and (not existing_title or existing_title == "查無資料"):
            if not brand_name or brand_name == "查無資料":
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
            
            print(f"✅ 完成回填: {official_title} | {phone}")
            time.sleep(1.5) # 稍微加長間隔，確保搜尋穩定

    print("🏁 所有品牌處理完畢！")

if __name__ == "__main__":
    main()
