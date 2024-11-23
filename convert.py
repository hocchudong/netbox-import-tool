import pandas as pd

# Đường dẫn tới file Excel
file_path = 'test.xlsx'

# Đọc dữ liệu từ sheet "Input"
df = pd.read_excel(file_path, sheet_name='Input')
df.columns = df.columns.str.strip()  # Xóa khoảng trắng ở tên cột

# Kiểm tra xem cột Rack có giá trị hay không
if df['Rack'].dropna().empty:
    print("Cột 'Rack' không có giá trị. Dừng xử lý.")
else:
    # Hàm xác định vai trò dựa trên giá trị cột Rack
    def get_role(rack_value):
        if isinstance(rack_value, str):  # Kiểm tra giá trị không phải NaN và là chuỗi
            if rack_value.startswith('FW'):
                return 'Firewall'
            elif rack_value.startswith('SW'):
                return 'Switch'
            elif rack_value.startswith('SRV'):
                return 'Server'
            elif rack_value.startswith('R'):
                return 'Router'
            else:
                return 'Unknown'
        return None  # Trả về None nếu giá trị trống

    # Thêm cột `role` vào DataFrame
    df['role'] = df['Rack'].apply(get_role)

    # Lọc bỏ các dòng có giá trị trống ở cột `rack` và `role`
    df = df.dropna(subset=['Rack', 'role'])

    # Định nghĩa các cột đầu ra cần có
    output_columns = ['role', 'manufacturer', 'device_type', 'status', 'site', 'name', 'tenant', 'serial', 'rack', 'position', 'face']

    # Chuẩn bị DataFrame đầu ra
    df_csv = pd.DataFrame(columns=output_columns)

    # Lấp dữ liệu từ DataFrame gốc vào các cột tương ứng
    df_csv['role'] = df['role']
    df_csv['manufacturer'] = df['Manufacturer']
    df_csv['device_type'] = df['Device Types']
    df_csv['serial'] = df['Serial Number']
    df_csv['name'] = df['Rack']
    df_csv['position'] = df['Position']

    # Gán giá trị mặc định cho các cột còn lại
    df_csv['status'] = 'active'      # Trạng thái mặc định
    df_csv['site'] = 'VNPT NTL'      # Site mặc định
    df_csv['rack'] = 'Rack 01'       # Tên mặc định
    df_csv['tenant'] = 'HTV'         # Tenant mặc định
    df_csv['face'] = 'front'         # Face mặc định

    # Lưu dữ liệu ra file CSV
    output_file_path = 'output_file1.csv'
    df_csv.to_csv(output_file_path, index=False)

    print(f"File CSV đã được lưu thành công tại: {output_file_path}")