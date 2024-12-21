# Nhập dữ liệu file
filepath = '/opt/netbox/netbox/plugin/netbox-import-tool/Auto_Import_Device/Version_4/Rack_M1-10.xlsx'
# filepath = '/opt/netbox/netbox/plugin/netbox-import-tool/Auto_Import_Device/Version_4/test_input.xlsx'
sheetname = 'Input'

# Nhập dữ liệu để kết nối tới NetBox
NetBox_URL = 'http://172.16.66.177:8000'
NetBox_Token = '633a7508b878bcbf33091699289a8a3026a3fbf6'

# Nhập tên Site của bạn
sitename = 'VNPT NTL'
status = 'active'

# Thông tin nếu add Rack Mới
width=19
u_height=42