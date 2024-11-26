import pandas as pd
from datetime import datetime
import config

# Đường dẫn tới file Excel
file_path = config.file_path

# Đọc dữ liệu từ sheet
df = pd.read_excel(file_path, sheet_name=config.sheet_name)
df.columns = df.columns.str.strip()  # Xóa khoảng trắng ở tên cột

# Kiểm tra xem cột Rack có giá trị hay không
if df['Name'].dropna().empty:
    print("Cột 'Name' không có giá trị. Dừng xử lý.")
else:
    # Hàm xác định vai trò dựa trên giá trị cột Rack
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
        # Kiểm tra nếu date_str là kiểu Timestamp, chuyển nó thành chuỗi trước
        if isinstance(date_str, pd.Timestamp):
            date_str = date_str.strftime('%d/%m/%Y')
        # Chuyển đổi chuỗi thành đối tượng datetime
        date_obj = datetime.strptime(date_str, '%d/%m/%Y')
        # Chuyển đối tượng datetime thành chuỗi với định dạng mới
        return date_obj.strftime('%Y-%m-%d')

    # Hàm xử lý trùng tên
    def handle_duplicate_names(df, name_col, rack_col, position_col):
        # Đếm tần suất xuất hiện của các giá trị trong cột Name
        name_counts = df[name_col].value_counts()
        duplicates = name_counts[name_counts > 1].index  # Lấy các giá trị Name bị trùng

        # Xử lý từng giá trị bị trùng
        for name in duplicates:
            duplicate_rows = df[df[name_col] == name]  # Lọc các dòng trùng tên
            for idx, row in enumerate(duplicate_rows.index):
                rack_value = df.at[row, rack_col]  # Giá trị Rack
                position_value = df.at[row, position_col]  # Giá trị số U
                # Gán tên mới với thông tin Rack và Position
                df.at[row, name_col] = f"{name}_{rack_value}_U{position_value}"

        return df

    # Thêm cột `role` vào DataFrame
    df['role'] = df['Role'].apply(get_role)
    # Thêm cột `year of investment` vào DataFrame
    df['cf_year_of_investment'] = df['Year of Investment'].dropna().apply(date_transfer)
    # Lọc bỏ các dòng có giá trị trống ở cột `Role` và `role`
    df = df.dropna(subset=['Role', 'role'])

    # Xử lý trùng tên trong cột Name
    df = handle_duplicate_names(df, name_col='Name', rack_col='Rack', position_col='U')

    # Định nghĩa các cột đầu ra cần có
    output_columns = [
        'role', 'manufacturer', 'device_type', 'status', 'site', 'name',
        'serial', 'rack', 'position', 'face', 'comments',
        'cf_contract_number', 'cf_year_of_investment'
    ]

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
    df_csv['status'] = config.status    # Trạng thái mặc định
    df_csv['site'] = config.site     # Site mặc định
    df_csv['face'] = 'front'         # Face mặc định

    # Lưu dữ liệu ra file CSV với tên đầu ra theo yêu cầu
    current_time = datetime.now().strftime('%H%M%S_%d%m%Y')
    output_file_path = f"output_{df['Rack'][0]}_{current_time}.csv"
    df_csv.to_csv(output_file_path, index=False)

    print(f"File CSV đã được lưu thành công tại: {output_file_path}")
