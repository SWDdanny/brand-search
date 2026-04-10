import os
import json
import time
import re
import requests
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googlesearch import search

# --- 設定區 ---
SPREADSHEET_ID = '1jb7MZ5w00zNs3T_I7lxT24nEChudAUnUnpXLm77sOXU'
SHEET_NAME = '品牌名單' 
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_service():
    service_account_info = os.environ.get("GCP_SERVICE_ACCOUNT")
    if service_account_info:
        info = json.loads(service_account_info)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = service_account.Credentials.from_service_account_file('service_account.json', scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def crawl_twincn(brand_name):
    query = f"{brand_name} 台灣公司網 twincn"
    print(f"🔍 搜尋中: {query}")
    try:
        # 增加 pause 避免被封鎖
        urls = search(query, num_results=5, lang="zh-TW", pause=5.0)
        for url in urls:
            if "twincn.com/item.aspx" in url or "twincn.com/L_item.aspx" in url:
                headers = {'User-Agent': 'Mozilla/5.0'}
                resp = requests.get(url, headers=headers, timeout=15)
                resp.encoding = 'utf-8'
                soup = BeautifulSoup(resp.text, 'html.parser')
                
                # 抓取抬頭
                h1_tag = soup.find('h1')
                company_title = h1_tag.text.strip().replace("公司基本資料", "") if h1_tag else ""
                
                # 抓取電話
                phone = ""
                phone_match = re.search(r'0\d{1,2}-\d{6,8}', soup.get_text())
                if phone_match: phone = phone_match.group()
                
                return company_title, phone
        return None, None
    except Exception as e:
        print(f"❌ 搜尋出錯: {e}")
        return None, None

def main():
    service = get_service()
    sheet = service.spreadsheets()
    
    # 讀取整張表
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{SHEET_NAME}!A:K").execute()
    values = result.get('values', [])
    if not values: return

    header = values[0]
    # 確保抓到正確欄位索引
    try:
        col_brand = header.index("品牌名稱")
        col_status = header.index("狀態")
    except ValueError:
        print("❌ 找不到欄位，請確認標題列是否有 '品牌名稱' 與 '狀態'")
        return

    for i, row in enumerate(values):
        if i == 0: continue # 跳過標題列
        if len(row) <= col_status: continue
        
        brand_name = row[col_brand]
        status = row[col_status]
        
        # 核心防護：只有「已分配」且「J欄(索引9)沒有資料」才執行
        # 且必須確保這一行真的有品牌名稱
        if status == "已分配" and brand_name:
            if len(row) <= 9 or not row[9]: # J欄為空才執行
                print(f"🚀 處理品牌: {brand_name}")
                official_title, phone = crawl_twincn(brand_name)
                
                # 第二道防護：如果抓不到東西，跳過，不進行任何 Update 動作
                if not official_title and not phone:
                    print(f"⏭️ 抓不到 {brand_name} 的資料，保留原狀，不更新。")
                    continue
                
                # 準備更新 (只針對 J 與 K 欄)
                update_range = f"{SHEET_NAME}!J{i+1}:K{i+1}"
                body = {'values': [[official_title or "查無抬頭", phone or "查無電話"]]}
                
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()
                print(f"✅ 更新成功: {official_title}")
                time.sleep(10) # 休息避免被擋

if __name__ == "__main__":
    main()
