import os
from boxsdk import Client, OAuth2
import time

CLIENT_ID = 'ljhnkofkfv6mw4jgrwxunett17il1ynq'
CLIENT_SECRET = 'DihXeH0t9snZYruJmpSuJjOL9KGRx8u7'
ACCESS_TOKEN = 'HtX0R3Vu26bzRdTRdYF5Nc3FFOFbSdH7'

# 設定 OAuth2
oauth2 = OAuth2(client_id=CLIENT_ID, client_secret=CLIENT_SECRET, access_token=ACCESS_TOKEN)
client = Client(oauth2)

# 要上傳的資料夾路徑和 Box 資料夾 ID
local_folder_path = r'C:\Users\200994\Downloads\000-BAT工具箱\Python\dist'
box_folder_id = '0'  # 根目錄，或替換為特定的 Box 資料夾 ID

# 遍歷資料夾中的所有檔案並上傳
for filename in os.listdir(local_folder_path):
    file_path = os.path.join(local_folder_path, filename)  # 完整的檔案路徑

    # 檢查是否為檔案（而不是資料夾）
    if os.path.isfile(file_path):
        print(f'Uploading {file_path} to Box...')
        client.folder(box_folder_id).upload(file_path, file_name=filename)

    # 重試邏輯，以便在收到連接錯誤時自動重新發送請求。這样可以在遇到暫時的問題時提高穩定性
    for retry in range(3):
       try:
           client.folder(box_folder_id).upload(file_path, file_name=filename)
           break  # 上傳成功，跳出重試循環
       except Exception as e:
           print(f"Upload failed: {e}. Retrying in 5 seconds...")
           time.sleep(5)  # 等待5秒再重試

print("All files uploaded successfully!")