# Hướng dẫn Cài đặt và Chạy Ứng dụng Quản lý Sửa chữa Thiết bị Y tế (Local - Dữ liệu JSON - Truy cập Mạng Nội Bộ)

Đây là hướng dẫn để bạn có thể chạy ứng dụng quản lý yêu cầu sửa chữa thiết bị y tế trên máy tính cá nhân của mình và cho phép các máy tính khác trong cùng mạng nội bộ truy cập. Phiên bản này lưu trữ dữ liệu trong các file JSON ở phía backend, và backend cũng sẽ phục vụ các file HTML của ứng dụng.

## Yêu cầu Hệ thống

1.  Python Phiên bản 3.7 trở lên.
2.  pip Trình quản lý gói của Python.
3.  Trình duyệt Web Chrome, Firefox, Edge, Safari phiên bản mới nhất.
4.  File Mẫu DOCX
     `Phieu_Yeu_Cau_Sua_Chua.docx`
     `Bien_Ban_Kiem_Tra (4).docx`
    (Cần được đặt trong thư mục `backendtemplates_docx` và chứa các thẻ Jinja2)
5.  Mạng Nội bộ (LANWLAN) Tất cả các máy tính (máy chủ chạy backend và các máy client truy cập) phải kết nối cùng một mạng.

## Cấu trúc Thư mục Bắt buộc

Để ứng dụng hoạt động đúng cách khi backend phục vụ HTML, cấu trúc thư mục của bạn phải như sau


medical_repair_app
├── backend
│   ├── app.py                          # File Python backend (Flask)
│   ├── served_html                    # THƯ MỤC CHỨA CÁC FILE HTML ĐỂ FLASK PHỤC VỤ
│   │   ├── admin_app.html              # File HTML cho Admin (BACKEND_URL phải là 'api')
│   │   └── department_app.html         # File HTML cho KhoaPhòng (BACKEND_URL phải là 'api')
│   ├── templates_docx                 # Thư mục chứa các file DOCX mẫu
│   │   ├── Phieu_Yeu_Cau_Sua_Chua.docx
│   │   └── Bien_Ban_Kiem_Tra (4).docx
│   ├── users.json                      # File lưu dữ liệu người dùng
│   ├── departments.json                # File lưu dữ liệu khoaphòng
│   ├── repair_requests.json            # File lưu dữ liệu yêu cầu sửa chữa
│   └── venv                           # (Môi trường ảo - tùy chọn)
└── README.md                           # File hướng dẫn này

Lưu ý quan trọng về file HTML trong `served_html`
 Các file `admin_app.html` và `department_app.html` đặt trong thư mục `backendserved_html` phải có hằng số `BACKEND_URL` được đặt là `'api'` trong phần JavaScript của chúng. Ví dụ
    ```javascript
    const BACKEND_URL = 'api'; 
    ```
    Điều này cho phép JavaScript tự động gửi yêu cầu API đến đúng máy chủ đã phục vụ file HTML.

## Các bước Cài đặt và Chạy

### 1. Cài đặt Backend (Python Flask) - Trên Máy Chủ

Chọn một máy tính trong mạng nội bộ để làm máy chủ chạy backend.

a.  Di chuyển đến thư mục `backend`
    ```bash
    cd pathtoyourmedical_repair_appbackend
    ```

b.  (Khuyến khích) Tạo và kích hoạt môi trường ảo
    ```bash
    python -m venv venv
    ```
     Windows `venvScriptsactivate`
     macOSLinux `source venvbinactivate`

c.  Cài đặt các thư viện Python cần thiết
    ```bash
    python -m pip install --upgrade pip
    pip install Flask Flask-CORS python-docx-template python-docx
    ```

d.  Chuẩn bị file mẫu DOCX Đặt vào `backendtemplates_docx` và đã chèn thẻ Jinja2.

e.  Chuẩn bị file HTML cho Backend phục vụ
     Tạo thư mục `served_html` bên trong thư mục `backend` nếu nó chưa tồn tại.
     Sao chép file `admin_app.html` và `department_app.html` của bạn vào thư mục `backendserved_html`.
     Đảm bảo rằng trong hai file HTML này, hằng số `BACKEND_URL` trong phần JavaScript đã được đặt thành `'api'`.

f.  Chạy Backend Server (Cho phép truy cập từ mạng nội bộ)
    Mở file `app.py` và đảm bảo dòng cuối cùng là
    ```python
    if __name__ == '__main__'
        app.run(host='0.0.0.0', port=5000, debug=True)
    ```
    Sau đó, chạy server
    ```bash
    python app.py
    ```
    Server sẽ lắng nghe trên tất cả các giao diện mạng của máy chủ, cổng 5000. Giữ cửa sổ Terminal này mở. Nó cũng sẽ phục vụ các file HTML từ thư mục `served_html`.

### 2. Xác định Địa chỉ IP Nội bộ của Máy Chủ Backend

Trên máy tính đang chạy backend `app.py`, bạn cần tìm địa chỉ IP nội bộ của nó (ví dụ `192.168.1.100`). Cách thực hiện
 Windows `ipconfig` trong Command Prompt.
 macOSLinux `hostname -I` hoặc `ip addr show` trong Terminal.

### 3. Cấu hình Tường lửa (Firewall) - Trên Máy Chủ Backend

Nếu máy chủ có tường lửa, bạn cần tạo một quy tắc cho phép các kết nối đến (inbound) trên cổng `5000`.

### 4. Truy cập Ứng dụng từ Trình duyệt

Từ bất kỳ máy tính nào trong cùng mạng nội bộ (bao gồm cả máy chủ), mở trình duyệt web và truy cập

 Để vào trang Admin
    `httpĐỊA_CHỈ_IP_MÁY_CHỦ5000`
    (Ví dụ `http192.168.1.1005000`)

 Để vào trang KhoaPhòng
    `httpĐỊA_CHỈ_IP_MÁY_CHỦ5000department`
    (Ví dụ `http192.168.1.1005000department`)

    Nếu truy cập từ chính máy chủ, bạn cũng có thể dùng `httplocalhost5000` hoặc `httplocalhost5000department`.

### 5. Sử dụng Ứng dụng

 Người dùng giờ đây có thể sử dụng ứng dụng thông qua các địa chỉ web trên.
 Dữ liệu sẽ được gửi đến và lấy từ máy chủ backend qua mạng nội bộ, và được lưu trữ trong các file JSON trên máy chủ.

## Xử lý Sự cố Cơ bản

 Lỗi Not Found khi truy cập `httpIP_SERVER5000`
     Đảm bảo `admin_app.html` (với `BACKEND_URL = 'api';`) nằm trong thư mục `backendserved_html`.
     Đảm bảo `department_app.html` (với `BACKEND_URL = 'api';`) nằm trong thư mục `backendserved_html`.
     Kiểm tra xem backend `app.py` có đang chạy và không có lỗi khởi động nào không.
 Không truy cập được từ máy client (nhưng truy cập được từ `localhost` trên máy chủ)
     Kiểm tra địa chỉ IP bạn nhập vào trình duyệt có chính xác không.
     Kiểm tra tường lửa trên máy chủ Đảm bảo tường lửa cho phép kết nối đến cổng 5000.
     Kiểm tra kết nối mạng giữa máy client và máy chủ (ví dụ thử `ping địa_chỉ_ip_máy_chủ`).
 Lỗi Failed to fetch trong console của trình duyệt
     Backend `app.py` có thể không chạy, hoặc có lỗi khi xử lý API. Kiểm tra Terminal của backend.
     Nếu bạn đã chỉnh sửa `BACKEND_URL` trong các file HTML trong `served_html` thành một địa chỉ IP cụ thể thay vì `'api'`, đó có thể là nguyên nhân. Khi Flask phục vụ HTML, `'api'` sẽ tự động trỏ đúng.
 Các lỗi khác Xem console của trình duyệt (F12) trên máy client và Terminal của backend trên máy chủ.

Chúc bạn thành công!
