# netbox-convert-tool
Công cụ chuyển đổi dữ liệu từ file xlsx thành csv để có thể bulk import vào NetBox

# Hướng dẫn sử dụng
Tool được sử dụng để lấy dữ liệu từ file `xlsx` với 1 mẫu định dạng đầu vào cố định. Sau đó sẽ thực hiện chuyển thành định dạng csv để bạn có thể dễ dàng *bulk import* lên hệ thống NetBox của mình

Có thể sử dụng đa nên tảng như *Windows*, *Linux*, *MacOS* một cách dễ dàng với **Python 3.0+**

Có 2 phiên bản để bạn có thể tùy chọn sử dụng với hệ thống của mình. Tuy nhiên, để phù hợp với đa số các hệ thống, tôi khuyên bạn nên chọn Version1
- Đối với Version1, phiên bản này không có các mục CustomField nên các bạn có thể trực tiếp sử dụng mà không cần thêm gì trên NetBox
- Đối với Version2, phiên bản này sẽ thêm vào 2 mục không có trong NetBox là *năm đầu tư* và *số hợp đồng*. Để có thể sử dụng được tool, các bạn sẽ cần thêm vào NetBox 2 mục nằm trong CustomField như sau:

![](/Anh/Screenshot_987.png)

## Hướng dẫn cài đặt
### Đối với Windows:
Tải xuống file code tại [đây](https://github.com/hocchudong/netbox-import-tool/blob/main/Conver%20XLSX%20to%20CSV/Version%201/device_convert.py)

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
Tiến hành sửa trong file `config.py`

```
file_path = '<Nhập đường dẫn tới file xlsx của bạn>'

sheet_name='<Nhập tên Sheet chứa data>'

output_file_path = '<Nhập đường dẫn lưu file đầu ra>'
```
## Thực thi chương trình
Sau khi hoàn tất chỉnh sửa file config, các bạn khởi chạy code.
- Đối với Windows các bạn sử dụng VSCode hoặc Python để chạy trực tiếp
- Đối với Linux:
```
chmod +x device_convert_new.py # cấp quyền khởi chạy

.device_convert_new.py         # Khởi chạy
```

Mẫu đầu ra khi khởi chạy thành công chính là file csv của bạn

Ví dụ: Khi bạn có 1 Input như file (`sample_input_for_convert.xlsx`)[sample_input_for_convert.xlsx] 

![](/Anh/Screenshot_988.png)

Bạn sẽ nhận được file `output_M04_233440_20241126.csv` như sau:

![](/Anh/Screenshot_989.png)

