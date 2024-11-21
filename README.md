# netbox-import-tool
Công cụ import dữ liệu vào netbox tự động

# Hướng dẫn sử dụng
## Config
Trước tiên, hãy cài đặt các thư viện yêu cầu trong `requirement.txt`

Tiến hành sửa code

```
file_path = 'Nhập đường dẫn tới file xlsx'
df = pd.read_excel(file_path, sheet_name='Nhập vào tên sheet')

NETBOX_URL = '<Nhập vào URL netbox>'
NETBOX_TOKEN = 'Nhập vào Token Netbox'
```
## Khởi chạy
Sau khi khởi chạy, Tool sẽ đưa ra danh sách các devicetype hiện có trong file xlsx mà không có trong NetBox. Công việc của bạn lúc này là lựa chọn việc mà bạn muốn làm với các dữ liệu còn thiếu đó.
- Chọn `1` để thêm thủ công thông qua giao diện Web
- Chọn `2` để thêm đơn giản với các thông tin nhập vào
- Chọn `3` để thêm tự động với các thông tin cơ bản.

