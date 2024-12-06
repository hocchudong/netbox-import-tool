from datetime import datetime
# Nhập đường dẫn tới file
file_path = 'sample_input.xlsx'

# Nhập một số thông tin về NetBox của bạn
status = 'active'

site = 'IDC NTL'

face = 'front'

# Đường dẫn lưu file
current_time = datetime.now().strftime('%H%M%S_%d%m%Y')
output_file_path = f"output_{current_time}.csv"