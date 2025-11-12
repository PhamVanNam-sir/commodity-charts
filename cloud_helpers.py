from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import git
import datetime
import os


def authenticate():
    """
    Xác thực với Google Drive.
    """
    gauth = GoogleAuth() # 1. Tạo đối tượng
    
    # 2. THIẾT LẬP QUYỀN 'OFFLINE' TRƯỚC KHI TẢI
    # Đây là cách chính xác để yêu cầu refresh_token
    gauth.settings['oauth_scope'] = ['https://www.googleapis.com/auth/drive']
    gauth.settings['access_type'] = 'offline'
    gauth.settings['approval_prompt'] = 'force' # Luôn luôn hỏi lại (để đảm bảo)

    # 3. Thử tải credentials đã lưu
    try:
        gauth.LoadCredentialsFile("credentials.json")
    except FileNotFoundError:
        pass # Không sao nếu không tìm thấy

    # 4. Kiểm tra
    if gauth.credentials is None:
        # 5. Đăng nhập (KHÔNG có 'auth_params' nữa)
        print("  Không tìm thấy credentials, mở trình duyệt để xác thực...")
        gauth.LocalWebserverAuth() # <--- ĐÂY LÀ DÒNG GỐC (đã bỏ auth_params)
        
    elif gauth.access_token_expired:
        # 6. Làm mới nếu hết hạn
        print("  Access token hết hạn, đang tự động làm mới...")
        gauth.Refresh()
    else:
        # 7. Đã xác thực
        print("  Đã có thông tin xác thực, đang ủy quyền...")
        gauth.Authorize()
        
    # 8. Lưu lại (quan trọng)
    print("  Lưu credentials vào file 'credentials.json'...")
    gauth.SaveCredentialsFile("credentials.json")
    
    return GoogleDrive(gauth)

def upload_or_update_file(drive, local_path, file_name, folder_id):
    """
    Hàm upload/update chính (dùng cho file Excel).
    Tìm kiếm, nếu có thì update, không thì tạo mới.
    TRẢ VỀ file object của Google Drive sau khi upload.
    """
    
    # Kiểm tra xem file local có tồn tại không
    if not os.path.exists(local_path):
        print(f"LỖI: File local không tồn tại: {local_path}")
        return None

    # 1. Tạo truy vấn tìm kiếm file
    query = f"title='{file_name}' and '{folder_id}' in parents and trashed=false"
    
    # 2. Thực hiện tìm kiếm
    file_list = drive.ListFile({'q': query}).GetList()
    
    drive_file = None
    
    if len(file_list) > 0:
        # --- Đã tìm thấy file -> CẬP NHẬT ---
        file_id = file_list[0]['id']
        print(f"  Đã tìm thấy file Drive. Đang cập nhật: {file_name} (ID: {file_id})")
        drive_file = drive.CreateFile({'id': file_id})
        
    else:
        # --- Không tìm thấy file -> UPLOAD MỚI ---
        print(f"  Không tìm thấy file Drive. Đang upload file mới: {file_name}")
        drive_file = drive.CreateFile({
            'title': file_name,
            'parents': [{'id': folder_id}]
        })
    
    # Set nội dung và Upload
    drive_file.SetContentFile(local_path)
    drive_file.Upload()
    print(f"  Upload/Cập nhật thành công: {file_name}")
    
    # Trả về file object
    return drive_file

def upload_html_and_get_link(drive, local_path, file_name, folder_id):
    """
    Hàm đặc biệt: Upload file HTML, set quyền public và trả về link web.
    """
    print(f"  Đang xử lý HTML: {file_name}")
    
    # 1. Upload hoặc Update file (dùng hàm chung)
    drive_file = upload_or_update_file(drive, local_path, file_name, folder_id)
    
    if drive_file is None:
        return None

    # 2. Set quyền "anyone with the link can view"
    # (Để tránh gọi API không cần thiết, kiểm tra xem đã public chưa)
    permissions = drive_file.GetPermissions()
    is_public = any(p['type'] == 'anyone' and p['role'] == 'reader' for p in permissions)
    
    if not is_public:
        print(f"    Set quyền public cho: {file_name}")
        drive_file.InsertPermission({
            'type': 'anyone',
            'value': 'anyone',
            'role': 'reader'
        })
    else:
        print(f"    File đã public, bỏ qua set quyền.")

    # 3. Trả về link để xem trên web (KHÔNG PHẢI link edit)
    # 'alternateLink' là link xem trên web (vd: drive.google.com/file/d/...)
    # 'webViewLink' cũng tương tự, dùng 'alternateLink' ổn định hơn
    file_id = drive_file['id']
    preview_link = f"https://drive.google.com/file/d/{file_id}/preview"
    
    print(f"    Tạo link preview thành công: {preview_link}")
    return preview_link

def get_or_create_folder(drive, folder_name, parent_folder_id):
    """
    Tìm kiếm một thư mục con theo tên bên trong thư mục cha.
    Nếu không thấy, tạo thư mục con mới.
    Trả về ID của thư mục con.
    """
    
    # 1. Tạo truy vấn tìm kiếm thư mục con
    # (Tìm chính xác tên, đúng thư mục cha, là thư mục, và không bị xóa)
    query = (f"title='{folder_name}' and "
             f"'{parent_folder_id}' in parents and "
             f"mimeType='application/vnd.google-apps.folder' and "
             f"trashed=false")
    
    file_list = drive.ListFile({'q': query}).GetList()
    
    if len(file_list) > 0:
        # --- Đã tìm thấy thư mục -> Trả về ID ---
        folder_id = file_list[0]['id']
        print(f"Đã tìm thấy thư mục con '{folder_name}' (ID: {folder_id})")
        return folder_id
    else:
        # --- Không tìm thấy -> Tạo mới ---
        print(f"Không tìm thấy thư mục con '{folder_name}'. Đang tạo mới...")
        folder_metadata = {
            'title': folder_name,
            'parents': [{'id': parent_folder_id}], # Đặt thư mục cha
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = drive.CreateFile(folder_metadata)
        folder.Upload() # Upload để tạo
        new_folder_id = folder['id']
        print(f"Đã tạo thư mục con thành công (ID: {new_folder_id})")
        return new_folder_id
    
def push_to_github(repo_local_path, github_token, github_username, github_repo_name, commit_message=None):
    """
    Tự động add, commit, và push các thay đổi trong thư mục repo local lên GitHub.
    Sử dụng Token để xác thực.
    """
    if commit_message is None:
        commit_message = f"Auto-update charts {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    
    try:
        # 1. Mở repo local
        repo = git.Repo(repo_local_path)
        print(f"  Mở repo thành công tại: {repo_local_path}")

        # 2. Kiểm tra xem có thay đổi không
        if not repo.is_dirty(untracked_files=True):
            print("  Không có thay đổi nào trong repo. Bỏ qua push.")
            return True

        # 3. Thêm tất cả các file (mới hoặc đã sửa)
        print("  Đang thêm (add) tất cả các thay đổi...")
        repo.git.add(A=True)
        
        # 4. Commit
        print(f"  Đang commit với message: '{commit_message}'")
        repo.index.commit(commit_message)
        
        # 5. Push lên GitHub
        print("  Đang push lên GitHub...")
        # Tạo URL xác thực (https://<token>@github.com/<username>/<repo_name>.git)
        remote_url = f"https://{github_token}@github.com/{github_username}/{github_repo_name}.git"
        
        # Lấy remote 'origin' và set URL mới (để xác thực)
        origin = repo.remote(name='origin')
        origin.set_url(remote_url)
        
        # Push
        origin.push()
        
        print("  Push lên GitHub thành công!")
        return True
        
    except Exception as e:
        print(f"LỖI khi push lên GitHub: {e}")
        print("  Kiểm tra lại đường dẫn repo, token và tên username/repo.")
        return False