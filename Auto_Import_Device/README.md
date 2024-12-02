# Auto-Import-Device
Tool được sử dụng để tự động Import Device lên NetBox với dữ liệu lấy từ file Excel theo 1 form đầu vào.

## Một vài thông tin về Tool
Tool sẽ thực hiện đọc dữ liệu từ file Excel của bạn, sau đó sẽ tự động thực hiện một số kiểm tra. 
- Kiểm tra tồn tại file, kiểm tra kết nối tới NetBox
- Kiểm tra các trường dữ liệu đã có đầy đủ trên NetBox hay chưa
- Riêng đối với phần Device-Types sẽ có thêm phần tự động Add những Device-Types chưa tồn tại trên NetBox

Tool có thể chạy trên đa hệ điều hành: Windows, Linux, MacOS, yêu cầu phiên bản ***Python 3.7*** trở lên và ***NetBox 4.0*** trở lên

**Một số yêu cầu cần làm trước khi khởi chạy Tools**
- Kiểm tra file đầu vào đã đúng với định dạng theo như file [`sample_input.xlsx`](https://github.com/hocchudong/netbox-import-tool/blob/main/Auto_Import_Device/sample_input.xlsx) hay chưa
- Kiểm tra các trường dữ liệu cần thiết trong file đã có đầy đủ trên NetBox hay chưa
- Kiểm tra file [`config.py`](https://github.com/hocchudong/netbox-import-tool/blob/main/Auto_Import_Device/config.py) xem đã cập nhật các biến cần thiết đúng hay chưa
- Kiểm tra xem đã cài đặt đầy đủ các gói ở trong file [`requirements.txt`](https://github.com/hocchudong/netbox-import-tool/blob/main/Auto_Import_Device/requirements.txt) hay chưa
- Kiểm tra xem bạn đã cấp quyền khởi chạy file hay sử dụng quyền sudo đối với Linux hay chưa.

## Hướng dẫn sử dụng
Hướng dẫn cài đặt và cấu hình để sử dụng tool
### Hướng dẫn cài đặt
**Windows**

- Các bạn tải xuống các file trong thư mục [`Auto_import_devices`](https://github.com/hocchudong/netbox-import-tool/blob/main/Auto_Import_Device). 
- Nên mở trong **Visual Studio Code** để đọc và kiểm soát file tốt hơn
- Mở Terminal và cài đặt các gói yêu cầu:
```
pip install -r requirements.txt
```
**Linux**
- Đối với cả CentOS và Ubuntu, các bạn dùng tổ hợp lệnh sau để thực hiện clone thư mục về:
```
# Clone git
git clone --depth 1 --filter=blob:none --sparse https://github.com/hocchudong/netbox-import-tool.git
```
```
# Giữ lại thư mục cần thiết
cd netbox-import-tool
git sparse-checkout set Auto_Import_Device
```
- Kích hoạt môi trường ảo và cài đặt các gói yêu cầu:
```
pip install -r requirements.txt
```
### Hướng dẫn cấu hình
Mở file `config.py` và thực hiện cấu hình như sau
```
# Nhập dữ liệu file
filepath = '<Nhập đường dẫn tới file xlsx của bạn>'
sheetname = '<Nhập tên Sheet chứa data>'

# Nhập dữ liệu để kết nối tới NetBox
NetBox_URL = '<Nhập URL NetBox>'
NetBox_Token = '<Nhập Token NetBox>'

# Nhập tên Site của bạn
sitename = '<Nhập tên Site>'
```
### Trước khi sử dụng Tool(chỉ dành cho Linux)
Sau khi đã cài đặt và hoàn thành các bước cấu hình, chúng ta cần làm:
- Bật quyền root
```
sudo su
```
- Cấp quyền khởi chạy file
```
chmod +x auto_import_devices.py
```
### Thực thi chương trình
**Windows**
- Các bạn có thể sử dụng ***CMD*** hoặc ***Visual Studio Code***,.. để chạy code

**Linux**
- Chạy file trong môi trường ảo
```
python auto_import_devices.py
```
- Hoặc sử dụng câu lệnh
```
.auto_import_devices
```
### Mẫu khi việc khởi chạy hoàn tất
Sau khi việc thực thi hoàn tất, mẫu kết quả sẽ như sau:

![](/Anh/Screenshot_990.png)

## Thêm
Nếu có bất kỳ trường hợp lỗi sử dụng, vui lòng góp ý ở mục Issue của GitHub

Dữ liệu của 2 cột trong mục ***CustomFields***: **Contract Number** và **Year of Investment** trên NetBox sẽ có định dạng như sau

![](/Anh/Screenshot_991.png)

Mặc định các DeviceTypes mới khi được khởi tạo tự động sẽ có số U là 1 và phiên bản 1.0 sẽ chỉ có thể làm việc với file dữ liệu có các Device Type là 1 dòng. 

Tool sẽ được cập nhật để lấy các cột Merge sớm nhất có thể!
