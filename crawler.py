import os
import requests
import json
import time
import re
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
import urllib.parse

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
        print(f"❌ Serper API 請求失敗: {e}")
        return []

def clean_company_name(raw_title):
    # 清理標題，相容如圓山大飯店等特殊抬頭
    name = re.sub(r'^(投標廠商|公司名稱|廠商名稱|公司抬頭|基本資料|公司簡介)[:：\s]+', '', raw_title)
    name = name.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    name = re.sub(r'[\(（].*?[\)）]', '', name).strip()
    name = re.sub(r'(台灣標案網|台灣公司網|104人力銀行|1111人力銀行|搜尋公司列表).*$', '', name).strip()
    return name

def extract_phone(text):
    """
    全格式電話提取：相容 (02)、空格、橫槓 -、分機 #
    """
    if not text: return None
    # 正則強化版
    phone_pattern = r'\(?0\d{1,2}\)?[\s-]?\d{3,4}[\s-]?\d{3,4}(?:\s?#\d+)?'
    match = re.search(phone_pattern, text)
    return match.group().strip() if match else None

def get_info_from_twincn_page(url):
    """
    進入內頁檢查營業狀態與電話
    """
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            page_text = soup.get_text(separator=' ')
            
            # 營業狀態過濾
            inactive_keywords = ["停業以外之非營業中", "廢止", "歇業", "解散"]
            if any(k in page_text for k in inactive_keywords):
                return "已停業", None
            
            return "營業中", extract_phone(page_text)
    except Exception as e:
        print(f"❌ 內頁連線失敗 {url}: {e}")
    return "連線失敗", None

def search_company_info(brand_name):
    print(f"🔎 正在查詢: {brand_name}")
    # 增加「公司」關鍵字以精準鎖定母公司資訊
    results = serper_request(f"{brand_name} twincn 公司")
    
    found_inactive = False
    
    if results:
        # 優先遍歷搜尋結果中的 twincn 連結
        for item in results:
            link = item.get("link", "")
            snippet = item.get("snippet", "")
            title = item.get("title", "")
            
            if "twincn.com/item.aspx?no=" in link:
                current_title = clean_company_name(title)
                
                # 策略：只要標題或摘要包含品牌前兩個字，或它是 Google 推薦的第一筆 twincn 連結
                keyword = brand_name[:2]
                if keyword in current_title or keyword in snippet or item.get("position") == 1:
                    
                    # 檢查摘要是否有停業字眼
                    if any(k in snippet for k in ["停業", "廢止", "歇業", "解散"]):
                        print(f"⚠️ 摘要顯示已停業: {current_title}")
                        found_inactive = True
                        continue
                    
                    # 優先從摘要抓電話 (最快)
                    s_phone = extract_phone(snippet)
                    if s_phone:
                        print(f"✨ 摘要直抓電話: {s_phone}")
                        return current_title, s_phone
                    
                    # 摘要沒電話，進內頁
                    print(f"🌐 進入內頁檢查: {link}")
                    status, p_phone = get_info_from_twincn_page(link)
                    if status == "營業中":
                        if p_phone: return current_title, p_phone
                        else: continue # 沒電話就找下一個 twincn 連結
                    elif status == "已停業":
                        found_inactive = True
                        continue

    return ("已停業" if found_inactive else "查無品牌"), "查無資料"

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

    if not rows: return

    for i, row in enumerate(rows):
        while len(row) < 11: row.append("")
        brand_name = row[2].strip()
        status = row[7].strip()
        existing_title = row[9].strip()
        
        # 僅處理已分配且未正確填寫的資料
        if status == "已分配" and (not existing_title or existing_title in ["查無品牌", "查無資料"]):
            if not brand_name: continue
            
            official_title, phone = search_company_info(brand_name)
            
            row_num = i + 2
            update_range = f"{SHEET_NAME}!J{row_num}:K{row_num}"
            try:
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="RAW",
                    body={"values": [[official_title, phone]]}
                ).execute()
                print(f"✅ 完成回填: {brand_name} -> {official_title} | {phone}")
            except Exception as e:
                print(f"❌ 更新錯誤: {e}")
            
            time.sleep(1.2) # 避免 API 頻率限制

if __name__ == "__main__":
    main()
