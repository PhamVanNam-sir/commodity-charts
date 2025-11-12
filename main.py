import yfinance as yf
import pandas as pd
import os
from yahoo_charts import create_commodity_charts
from sunsirs_charts import create_excel_with_charts
from cloud_helpers import push_to_github, authenticate, upload_or_update_file

comodity = ['HRC=F', # Hot Rolled Coil
        'CL=F',  # Crude Oil (WTI)
        'BZ=F',  # Brent Crude
        'NG=F',  # Natural Gas
        'RB=F',  # RBOB Gasoline
        'HO=F',  # Heating Oil
        'GC=F',  # Gold
        'SI=F',  # Silver
        'HG=F',  # Copper
        'PL=F',  # Platinum
        'PA=F',  # Palladium
        'ALI=F', # Aluminum
        'DX=F',  # Dollar Index
        ]

period = '2y'
part = []
for i in comodity:
    get = yf.Ticker(i)
    a = get.history(period=period)
    a.index.name = 'date'
    prices = a['Close']
    a['name'] = i
    part.append(a)

df = pd.concat(part)  


UPLOAD_FILES = True

# --- Cấu hình Google Drive ---
ROOT_FOLDER_ID = '1tAeJoC2BiHTV_mTC0KU11ngP7rdcBV-M'

# --- Cấu hình GitHub ---
GITHUB_USERNAME = "PhamVanNam-sir" 
GITHUB_REPO_NAME = "commodity-charts"
REPO_LOCAL_PATH = "."
GITHUB_TOKEN = os.getenv("API_TOKEN") 
if not GITHUB_TOKEN:
    raise ValueError("LỖI: Không tìm thấy GITHUB_API_TOKEN.")
GITHUB_PAGES_URL = f"https://{GITHUB_USERNAME}.github.io/{GITHUB_REPO_NAME}/"

# --- Cấu hình Local-Mode ---
# (Thư mục sẽ được dùng nếu UPLOAD_FILES = False)
LOCAL_HTML_FOLDER = 'charts_html_local' 

# -----------------------------------------------------------------
# -------------------- KẾT THÚC CẤU HÌNH -------------------------
# -----------------------------------------------------------------


try:
    # --- BƯỚC 1: TẠO FILE YAHOO ---
    print("\n--- BƯỚC 1: Bắt đầu tạo file Yahoo Finance ---")
    
    if UPLOAD_FILES:
        print("  Chế độ: UPLOAD. Sẽ lưu HTML vào repo local và dùng link GitHub.")
        create_commodity_charts(df, # Giả sử 'df' đã có từ cell trên
                                'commodity_charts.xlsx', 
                                period_years=1,
                                upload_mode=True, # <-- Bật
                                github_repo_local_path=REPO_LOCAL_PATH,
                                github_pages_url=GITHUB_PAGES_URL
                               )
    else:
        print("  Chế độ: LOCAL. Sẽ lưu HTML vào thư mục local.")
        create_commodity_charts(df, # Giả sử 'df' đã có
                                'commodity_charts.xlsx', 
                                period_years=1,
                                upload_mode=False, # <-- Tắt
                                local_html_folder=LOCAL_HTML_FOLDER
                               )
    
    # --- BƯỚC 2: TẠO FILE SUNSIRS (LOCAL) ---
    # (Hàm này luôn chạy, vì nó chỉ tạo file local)
    print("\n--- BƯỚC 2: Bắt đầu tạo file Sunsirs (local) ---")
    commodities_to_fetch_sunsirs = [
        'Coking coal', 'Fuel Oil', 'Gasoline', 'Diesel', 
        'Hot rolled coil', 'Iron ore'
    ]
    create_excel_with_charts(commodities_to_fetch_sunsirs, 
                             output_filename="Sunsirs_Charts.xlsx")
    
    # --- BƯỚC 3 & 4: UPLOAD (NẾU ĐƯỢC BẬT) ---
    if UPLOAD_FILES:
        # --- BƯỚC 3: PUSH GITHUB ---
        print("\n--- BƯỚC 3: [UPLOAD=True] Bắt đầu push các file HTML lên GitHub ---")
        push_to_github(repo_local_path=REPO_LOCAL_PATH,
                       github_token=GITHUB_TOKEN,
                       github_username=GITHUB_USERNAME,
                       github_repo_name=GITHUB_REPO_NAME)
        
        # --- BƯỚC 4: UPLOAD GOOGLE DRIVE ---
        print("\n--- BƯỚC 4: [UPLOAD=True] Bắt đầu upload 2 file Excel lên Google Drive ---")
        
        print("  Đang xác thực Google Drive...")
        drive_service = authenticate()
        print("  Xác thực Google Drive thành công!")
        
        file_list_to_upload = [
            {"local_path": "commodity_charts.xlsx", "drive_name": "commodity_charts.xlsx"},
            {"local_path": "Sunsirs_Charts.xlsx", "drive_name": "Sunsirs_Charts.xlsx"}
        ]

        for file_info in file_list_to_upload:
            print(f"--- Đang xử lý file Excel: {file_info['local_path']} ---")
            upload_or_update_file(drive_service, 
                                  file_info['local_path'], 
                                  file_info['drive_name'], 
                                  ROOT_FOLDER_ID 
                                 )
        
        print("\n✅✅✅ HOÀN TẤT TOÀN BỘ QUY TRÌNH (UPLOAD)! ✅✅✅")
    
    else:
        # --- BƯỚC 3 & 4 (BỊ TẮT) ---
        print("\n--- BƯỚC 3&4: [UPLOAD=False] Bỏ qua bước push GitHub và upload Google Drive.")
        print("\n✅✅✅ HOÀN TẤT (LOCAL)! ✅✅✅")
        print(f"Các file Excel và HTML đã được tạo/cập nhật trong thư mục local (thư mục HTML: '{LOCAL_HTML_FOLDER}').")

except Exception as e:
    print(f"ĐÃ XẢY RA LỖI NGHIÊM TRỌNG: {e}")
    import traceback
    traceback.print_exc()