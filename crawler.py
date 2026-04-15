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
    # 移除標題雜質，放寬限制以符合圓山大飯店等特殊抬頭
    name = re.sub(r'^(投標廠商|公司名稱|廠商名稱|公司抬頭|基本資料|公司簡介)[:：\s]+', '', raw_title)
    name = name.split(' - ')[0].split(' | ')[0].split('｜')[0].split(' : ')[0].strip()
    name = re.sub(r'[\(（].*?[\)）]', '', name).strip()
    name = re.sub(r'(台灣標案網|台灣公司網|104人力銀行|1111人力銀行|搜尋公司列表).*$', '', name).strip()
    return name

def extract_phone(text):
    """
    通用電話提取：支持括號, 空格, #分機, 橫槓 (-)
    """
    if not text: return None
    # 正則：支持 (02) 2886-8888 #123 等多種格式
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
            
            # 檢查是否為營業中
            inactive_keywords = ["停業以外之非營業中", "廢止", "歇業", "解散"]
            if any(k in page_text for k in inactive_keywords):
                return "已停業", None
            
            return "營業中", extract_phone(page_text)
    except Exception as e:
        print(f"❌ 進入頁面失敗 {url}: {e}")
    return "連線失敗", None

def search_company_info(brand_name):
    print(f"🔎 搜尋品牌: {brand_name}")
    results = serper_request(f"{brand_name} twincn")
    
    found_inactive = False
    
    if results:
        for item in results:
            link = item.get("link", "")
            snippet = item.get("snippet", "")
            title = item.get("title", "")
            
            # 確保是台灣公司網的內頁
            if "twincn.com/item.aspx?no=" in link:
                current_title = clean_company_name(title)
                
                # --- 策略 1: 優先檢查 Snippet (摘要) ---
                # 檢查摘要是否有停業字眼
                if any(k in snippet for k in ["停業", "廢止", "歇業", "解散"]):
                    print(f"⚠️ 摘要顯示已停業: {current_title}")
                    found_inactive = True
                    continue
                
                # 從摘要嘗試抓取電話
                snippet_phone = extract_phone(snippet)
                if snippet_phone:
                    print(f"✨ 從摘要直接抓到電話: {snippet_phone}")
                    return current_title, snippet_phone
                
                # --- 策略 2: 摘要沒電話，則進入內頁 ---
                print(f"🌐 摘要無電話，進入內頁檢查: {link}")
                status, page_phone = get_info_from_twincn_page(link)
                
                if status == "營業中":
                    # 內頁若有電話就回傳，沒有就繼續找下一個搜尋結果
                    if page_phone:
                        return current_title, page_phone
                    else:
                        print(f"ℹ️ 內頁無電話，嘗試下一個搜尋結果...")
                        continue
                elif status == "已停業":
                    found_inactive = True
                    continue

    # 最終結果回傳機制
    final_title = "已停業" if found_inactive else "查無品牌"
    return final_title, "查無資料"

def main():
    service = get_gspread_service()
    sheet = service.spreadsheets()
    range_to_read = f"{SHEET_NAME}!A2:K"
    
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_to_read).execute()
        rows = result.get('values', [])
    except Exception as e:
        print(f"❌ 讀取試算表失敗: {e}")
        return

    if not rows:
        return

    for i, row in enumerate(rows):
        while len(row) < 11: row.append("")
        brand_name = row[2].strip()      # C 欄
        status = row[7].strip()          # H 欄
        existing_title = row[9].strip()  # J 欄
        
        # 僅處理 狀態為「已分配」且 J 欄還沒有正確資料的列
        if status == "已分配" and (not existing_title or existing_title in ["查無品牌", "查無資料"]):
            if not brand_name: continue
            
            official_title, phone = search_company_info(brand_name)
            
            row_num = i + 2
            update_range = f"{SHEET_NAME}!J{row_num}:K{row_num}"
            update_body = {"values": [[official_title, phone]]}
            
            try:
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="RAW",
                    body=update_body
                ).execute()
                print(f"✅ 完成回填: {brand_name} -> {official_title} | {phone}")
            except Exception as e:
                print(f"❌ 更新失敗: {e}")
            
            time.sleep(1.2)

    print("🏁 所有品牌處理完畢。")

if __name__ == "__main__":
    main()
