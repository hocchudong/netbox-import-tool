import pandas as pd
from datetime import datetime
import config

# Đường dẫn tới file Excel
file_path = config.file_path

# Đọc dữ liệu từ sheet "Input"
df = pd.read_excel(file_path, sheet_name=config.sheet_name)
df.columns = df.columns.str.strip()  # Xóa khoảng trắng ở tên cột

# Kiểm tra xem cột Name có giá trị hay không
if df['Name'].dropna().empty:
    print("Error while processing! Please try again or contact me!")
else:
    # Hàm xác định vai trò dựa trên giá trị cột Role
    def get_role(role_value):
        if isinstance(role_value, str):  # Kiểm tra giá trị không phải NaN và là chuỗi
            if role_value.lower() == 'fw':
                return 'Firewall'
            elif role_value.lower() == 'sw':
                return 'Switch'
            elif role_value.lower() == 'svr':
                return 'Server'
            elif role_value.lower() == 'r':
                return 'Router'
            else:
                return 'Unknown'
        return None  # Trả về None nếu giá trị trống

    # Hàm xử lý dữ liệu cột date
    def date_transfer(date_str):
        if isinstance(date_str, pd.Timestamp):
            date_str = date_str.strftime('%d/%m/%Y')
        date_obj = datetime.strptime(date_str, '%d/%m/%Y')
        return date_obj.strftime('%Y-%m-%d')

    # Hàm xử lý dữ liệu cột Name khi bị trùng
    def handle_duplicate_names(df, name_col, rack_col, position_col):
        name_counts = df[name_col].value_counts()
        duplicates = name_counts[name_counts > 1].index  # Các giá trị bị trùng

        def make_unique_name(row):
            if row[name_col] in duplicates:
                return f"{row[name_col]}_{row[rack_col]}_{row[position_col]}"
            return row[name_col]

        return df.apply(make_unique_name, axis=1)

    # Hàm trích xuất tên Rack từ chuỗi
    def extract_rack_name(rack_value):
        if isinstance(rack_value, str) and "=" in rack_value:
            return rack_value.split("=")[-1]  # Lấy phần sau dấu "="
        return "Unknown"  # Giá trị mặc định nếu không hợp lệ

    # Xử lý cột Rack để lấy tên Rack
    df['Rack_Name'] = df['Rack'].apply(extract_rack_name)

    # Xử lý cột Name để đảm bảo không bị trùng
    df['Name'] = handle_duplicate_names(df, name_col='Name', rack_col='Rack_Name', position_col='U')

    # Thêm cột `role` vào DataFrame
    df['role'] = df['Role'].apply(get_role)
    # Thêm cột `year of investment` vào DataFrame
    df['cf_year_of_investment'] = df['Year of Investment'].dropna().apply(date_transfer)
    # Lọc bỏ các dòng có giá trị trống ở cột `Role` và `role`
    df = df.dropna(subset=['Role', 'role'])

    # Định nghĩa các cột đầu ra cần có
    output_columns = ['role', 'manufacturer', 'device_type', 'status', 'site', 'name', 'serial', 'rack', 'position', 'face', 'comments', 'cf_contract_number', 'cf_year_of_investment']

    # Chuẩn bị DataFrame đầu ra
    df_csv = pd.DataFrame(columns=output_columns)

    # Lấp dữ liệu từ DataFrame gốc vào các cột tương ứng
    df_csv['role'] = df['role']
    df_csv['manufacturer'] = df['Manufacturer']
    df_csv['device_type'] = df['Device Type']
    df_csv['serial'] = df['Serial Number']
    df_csv['name'] = df['Name']
    df_csv['position'] = df['U']
    df_csv['cf_year_of_investment'] = df['cf_year_of_investment']
    df_csv['comments'] = df['Comments']
    df_csv['cf_contract_number'] = df['Contract Number']
    df_csv['rack'] = df['Rack']

    # Gán giá trị mặc định cho các cột còn lại
    df_csv['status'] = config.status # Trạng thái mặc định
    df_csv['site'] = config.site     # Site mặc định
    df_csv['face'] = 'front'         # Face mặc định

    # Lấy giá trị thời gian hiện tại
    current_time = datetime.now().strftime("%H%M%S")
    current_date = datetime.now().strftime("%Y%m%d")

    # Lấy giá trị rack_name từ cột Rack_Name
    rack_name = df['Rack'].iloc[0]  

    # Đặt tên file với định dạng 
    output_file_name = f"output_{rack_name}_{current_time}_{current_date}.csv"
    df_csv.to_csv(output_file_name, index=False)

    print(f"Complete save: {output_file_name}")
