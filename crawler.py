import os, json, time, re, requests, urllib.parse
from bs4 import BeautifulSoup
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- 設定區 ---
SPREADSHEET_ID = '1jb7MZ5w00zNs3T_I7lxT24nEChudAUnUnpXLm77sOXU'
SHEET_NAME = '品牌名單' 
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_service():
    info_str = os.environ.get("GCP_SERVICE_ACCOUNT")
    if info_str:
        return build('sheets', 'v4', credentials=service_account.Credentials.from_service_account_info(json.loads(info_str), scopes=SCOPES))
    return build('sheets', 'v4', credentials=service_account.Credentials.from_service_account_file('service_account.json', scopes=SCOPES))

def search_twincn_directly(brand_name):
    """直接使用台灣公司網內建搜尋，並精準解析 HTML"""
    encoded_name = urllib.parse.quote(brand_name)
    search_url = f"https://www.twincn.com/L_search.aspx?q={encoded_name}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    print(f"DEBUG: 正在台灣公司網搜尋 -> {brand_name}")
    try:
        # 1. 第一層：搜尋結果頁
        resp = requests.get(search_url, headers=headers, timeout=15)
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        # 尋找第一個搜尋結果連結
        link = soup.select_one('td a[href^="item.aspx"]')
        if not link:
            print(f"DEBUG: 台灣公司網內找不到 {brand_name}")
            return None, None
            
        target_url = "https://www.twincn.com/" + link['href']
        print(f"DEBUG: 找到公司頁面 -> {target_url}")
        
        # 2. 第二層：公司詳細資料頁
        item_resp = requests.get(target_url, headers=headers, timeout=15)
        item_resp.encoding = 'utf-8'
        item_soup = BeautifulSoup(item_resp.text, 'html.parser')
        
        # 抓取抬頭 (移除不需要的後綴)
        h1 = item_soup.find('h1')
        title = h1.text.strip().replace("公司基本資料", "").strip() if h1 else brand_name
        
        # 抓取電話邏輯 (雙重保險)
        phone = "查無電話"
        
        # 策略 A: 從基本資料表格解析
        tables = item_soup.find_all('table')
        for table in tables:
            for tr in table.find_all('tr'):
                tds = tr.find_all('td')
                if len(tds) >= 2 and "電話" in tds[0].get_text():
                    raw_phone = tds[1].get_text(separator=" ").strip()
                    # 匹配格式如 02-23958399
                    match = re.search(r'\(?0\d{1,2}\)?-\d{6,9}', raw_phone)
                    if match:
                        phone = match.group()
                        break
            if phone != "查無電話": break
            
        # 策略 B: 如果表格沒抓到，從 Meta Description 提取 (針對你提供的 HTML 結構優化)
        if phone == "查無電話":
            meta_desc = item_soup.find('meta', attrs={'name': 'og:description'}) or \
                        item_soup.find('meta', attrs={'name': 'description'})
            if meta_desc and meta_desc.get('content'):
                # 抓取「電話:」後面的號碼
                meta_match = re.search(r'電話:([\d-]+)', meta_desc.get('content'))
                if meta_match:
                    phone = meta_match.group(1)
        
        return title, phone
        
    except Exception as e:
        print(f"DEBUG: 爬蟲解析過程出錯 -> {e}")
        return None, None

def main():
    print("🚀 程式啟動...")
    service = get_service()
    sheet = service.spreadsheets()
    
    try:
        # 1. 讀取資料
        print(f"📂 正在讀取工作表: {SHEET_NAME}")
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{SHEET_NAME}!A:K").execute()
        values = result.get('values', [])
    except Exception as e:
        print(f"❌ 讀取 Sheet 失敗！錯誤: {e}")
        return

    if not values:
        print("⚠️ Sheet 是空的。")
        return

    header = values[0]
    
    # 找到關鍵欄位索引
    col_brand, col_status = -1, -1
    for idx, col in enumerate(header):
        if "品牌名稱" in col: col_brand = idx
        if "狀態" in col: col_status = idx

    if col_brand == -1 or col_status == -1:
        print(f"❌ 找不到「品牌名稱」或「狀態」欄位！")
        return

    for i, row in enumerate(values):
        if i == 0: continue # 跳過標題
        
        # 取得當前行的狀態與品牌
        status = row[col_status] if len(row) > col_status else ""
        brand = row[col_brand] if len(row) > col_brand else ""
        
        # 判斷是否需要處理：狀態為「已分配」且 J 欄 (Index 9) 為空或查無資料
        # row[9] 對應 Excel 的 J 欄
        has_official_title = len(row) > 9 and row[9].strip() != "" and row[9] != "查無資料"
        
        if status == "已分配" and brand and not has_official_title:
            print(f"🔎 處理中: {brand} (第 {i+1} 行)")
            official_title, phone = search_twincn_directly(brand)
            
            if official_title:
                # 回填資料到 J 和 K 欄 (J=10, K=11)
                update_range = f"{SHEET_NAME}!J{i+1}:K{i+1}"
                body = {'values': [[official_title, phone]]}
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="USER_ENTERED",
                    body=body
                ).execute()
                print(f"✅ 已回填: {official_title} / {phone}")
            else:
                # 若搜尋失敗，填入查無資料避免重複處理
                update_range = f"{SHEET_NAME}!J{i+1}"
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="USER_ENTERED",
                    body={'values': [["查無資料"]]}
                ).execute()
            
            time.sleep(3) # 避免請求過快被封鎖

    print("🏁 程式執行完畢。")

if __name__ == "__main__":
    main()
