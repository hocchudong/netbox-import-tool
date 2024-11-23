# netbox-import-tool
Công cụ import dữ liệu vào netbox tự động

# Hướng dẫn sử dụng
Tool được sử dụng để kiểm tra các devicetypes hiện đang có trong 1 file xlsx nhưng không có trên NetBox.

Có thể sử dụng đa nên tảng như *Windows*, *Linux*, *MacOS* một cách dễ dàng với **Python 3.0+**
## Hướng dẫn cài đặt
### Đối với Windows:
Tải xuống file code tại [đây](https://github.com/hocchudong/netbox-import-tool)

Sử dụng Visual Studio Code hoặc trình soạn thảo code Python.

Mở Terminal, chạy câu lệnh sau để cài đặt các thư viện cần thiết:
```
pip install -r requirement.txt
```
### Đối với Linux
- Đối với các hệ điều hành tương tự Ubuntu:
```
sudo apt update && sudo apt upgrade

sudo apt install python3 python3-pip
```
- Đối với các hệ điều hành tương tự CentOS
```
sudo yum update && sudo yum upgrade

sudo apt install python3 python3-pip
```
- Tạo môi trường ảo và cài đặt các thư viện cần thiết
```
python3 -m venv env                     # Tạo môi trường ảo
source env/bin/activate                 # Kích hoạt môi trường
pip install -r requirements.txt         # Cài đặt môi trường
```

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

Mẫu đầu ra khi khởi chạy thành công

![](/Anh/Screenshot_986.png)