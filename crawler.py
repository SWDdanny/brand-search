import os
import requests
import json
import time
import re
from bs4 import BeautifulSoup
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
        print(f"❌ Serper API 請求失敗: {e}")
        return []

def clean_company_name(raw_title):
    """
    清理標題，移除雜質取得正式抬頭
    """
    name = re.sub(r'^(投標廠商|公司名稱|廠商名稱|公司抬頭|基本資料|公司簡介)[:：\s]+', '', raw_title)
    name = name.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    name = re.sub(r'[\(（].*?[\)）]', '', name).strip()
    name = re.sub(r'(台灣標案網|台灣公司網|104人力銀行|1111人力銀行).*$', '', name).strip()
    return name

def is_valid_company_name(name):
    keywords = ["公司", "集團", "行號", "有限", "工作室", "企業", "社企"]
    invalid_keywords = ["一家", "身處", "領域", "官網", "介紹", "新聞", "評價", "職缺", "徵才"]
    has_keyword = any(k in name for k in keywords)
    not_descriptive = not any(ik in name for ik in invalid_keywords)
    is_proper_length = 4 <= len(name) < 25
    return has_keyword and not_descriptive and is_proper_length

def get_phone_from_twincn_direct(official_title):
    """
    修正版：先搜尋公司名稱取得正確 ID，再進入內頁抓取電話
    """
    print(f"🌐 正在台灣公司網搜尋: {official_title}")
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'Referer': 'https://twincn.com/'
    }
    
    search_url = f"https://twincn.com/Search.aspx?q={urllib.parse.quote(official_title)}"
    phone_pattern = r'\(?0\d{1,2}\)?[\s-]?\d{3,4}[\s-]?\d{3,4}(?:\s?#\d+)?'
    
    try:
        # 第一步：發送搜尋請求
        resp = requests.get(search_url, headers=headers, timeout=15)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            
            # 尋找搜尋結果清單中的第一個連結 (通常是 /item.aspx?no=數字)
            # 台灣公司網的結果通常在 class="Titem" 或 <a> 標籤內
            first_link = None
            for a in soup.find_all('a', href=True):
                if 'item.aspx?no=' in a['href']:
                    first_link = "https://twincn.com/" + a['href'].lstrip('/')
                    break
            
            # 第二步：如果找到正式內頁連結，進入內頁
            target_url = first_link if first_link else search_url
            if first_link:
                print(f"🔗 找到正式內頁網址: {target_url}")
                resp = requests.get(target_url, headers=headers, timeout=15)
                soup = BeautifulSoup(resp.text, 'html.parser')

            # 第三步：抓取電話 (從整頁純文字找)
            page_text = soup.get_text(separator=' ')
            phone_matches = re.findall(phone_pattern, page_text)
            
            if phone_matches:
                for p in phone_matches:
                    clean_p = p.strip()
                    # 排除掉長度太短的錯誤號碼 (區碼+號碼至少 9 碼)
                    if len(clean_p) >= 9:
                        return clean_p
                        
    except Exception as e:
        print(f"❌ 爬取台灣公司網出錯: {e}")
    
    return "查無資料"

def search_company_info(brand_name):
    # --- 步驟 1: 查找品牌正式抬頭 (使用指定關鍵字) ---
    print(f"🔎 步驟 1: 查找品牌正式抬頭 -> 關鍵字: {brand_name} twincn")
    results_step1 = serper_request(f"{brand_name} twincn")
    
    official_title = ""

    if results_step1:
        for item in results_step1:
            temp_name = clean_company_name(item.get("title", ""))
            if is_valid_company_name(temp_name):
                official_title = temp_name
                print(f"🎯 識別到正式抬頭: {official_title}")
                break
    
    # 如果第一步沒找到，直接回傳查無品牌
    if not official_title:
        print(f"❌ 無法識別 {brand_name} 的正式抬頭")
        return "查無品牌", "查無資料"

    # --- 步驟 2: 直接到台灣公司網抓電話 (不透過 Serper) ---
    phone = get_phone_from_twincn_direct(official_title)

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
        
        # 僅處理 狀態為「已分配」且 J欄還沒有資料的列
        if status == "已分配" and not existing_title:
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
            
            print(f"✅ 完成回填: {brand_name} -> {official_title} | {phone}")
            time.sleep(1.0) # 稍作停頓

    print("🏁 所有處理已結束。")

if __name__ == "__main__":
    main()
