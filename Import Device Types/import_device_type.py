import pandas as pd
#import requests
import re
import pynetbox
import urllib3

# Đọc dữ liệu file xlsx để tìm ra list các device types
file_path = 'import_device_type.xlsx'
df = pd.read_excel(file_path, sheet_name='New form')

# Lấy ra danh sách các device types từ file xlsx
device_types_in_file = df['Device Types'].dropna().drop_duplicates().tolist() 

# Thiết lập thông tin kết nối tới NetBox
NETBOX_URL = 'https://www.netboxlab.local'
NETBOX_TOKEN = '94c41d00fafaaf2132ab3abe97d03e57e5183168'
# Kết nối tới NetBox
nb = pynetbox.api(NETBOX_URL, token=NETBOX_TOKEN)
nb.http_session.verify = False                                      
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Kiểm tra DeviceTypes trên NetBox
device_types_not_in_netbox = []
for device_type in device_types_in_file:
    search_result = nb.dcim.device_types.filter(model=device_type)
    if not search_result:
        device_types_not_in_netbox.append(device_type)
        
# In ra danh sách
print("DeviceTypes not in NetBox:")
print(device_types_not_in_netbox)

# Chọn cách mà bạn muốn làm với các device types chưa có trong NetBox
print("\nWhat do you want with the device_types not in NetBox")
print("1. Add manual and quit")
print("2. Add easy with sample information")
print("3. Automatic add")

choice = input("Enter your choice? (1, 2 or 3): ")

if choice == '1':
    print("\n Please Add Device Types manually!")
    exit()

elif choice == '2':
    for device_type in device_types_not_in_netbox:
        print(f"\nEnter information for {device_type}")
        manufacturer_name = input("Manufacturer: ").strip()
        u_height = int(input("U-height: ").strip())
        is_full_depth = input("Full dept(yes/no): ").strip().lower() == 'yes'

        try:
            # Kiểm tra xem manufacturer đã tồn tại trên NetBox chưa
            manufacturer_records = nb.dcim.manufacturers.filter(name=manufacturer_name)
            
            # Duyệt qua tất cả manufacturers trả về
            manufacturer = None
            for record in manufacturer_records:
                manufacturer = record
                break  # Lấy manufacturer đầu tiên tìm thấy

            if manufacturer:
                print(f"Using existing manufacturer: {manufacturer.name} (ID: {manufacturer.id})")
            else:
                # Tạo slug hợp lệ từ manufacturer_name
                manufacturer_slug = re.sub(r'[^a-z0-9-]', '-', manufacturer_name.lower()).strip('-')
                
                # Nếu không có manufacturer, tạo mới
                manufacturer = nb.dcim.manufacturers.create(
                    name=manufacturer_name,
                    slug=manufacturer_slug  # Sử dụng slug hợp lệ
                )
                print(f"Created new manufacturer: {manufacturer.name} (ID: {manufacturer.id})")

            # Tạo slug hợp lệ cho device type
            device_type_slug = re.sub(r'[^a-z0-9-]', '-', device_type.lower()).strip('-')

            # Tạo mới device type
            new_device_type = nb.dcim.device_types.create({
                'model': device_type,
                'slug': device_type_slug,  # Slug cho device type
                'manufacturer': manufacturer.id,  # ID của manufacturer
                'u_height': u_height,
                'is_full_depth': is_full_depth,
            })
            print(f"Complete Add: {device_type}")
        
        except Exception as e:
            print(f"Error while we try to add {device_type}: {e}")
            break
elif choice == "3":
    print("You choose to Add Device Automatic with sample information")
    print("Trying to Add Automatic...")
    
    for device_type in device_types_not_in_netbox:
        try:
            # Lấy thông tin từ file xlsx
            matching_rows = df[df['Device Types'] == device_type]
            if matching_rows.empty:
                print(f"Device type {device_type} not found in Excel. Skipping...")
                continue

            row = matching_rows.iloc[0]
            manufacturer_name = row['Manufacturer'].strip()
            u_height = int(row['U'])

            # Kiểm tra manufacturer tồn tại hoặc tạo mới
            # Kiểm tra xem manufacturer đã tồn tại trên NetBox chưa
            manufacturer_records = nb.dcim.manufacturers.filter(name=manufacturer_name)
            
            # Duyệt qua tất cả manufacturers trả về
            manufacturer = None
            for record in manufacturer_records:
                manufacturer = record
                break  # Lấy manufacturer đầu tiên tìm thấy

            if manufacturer:
                print(f"Using existing manufacturer: {manufacturer.name} (ID: {manufacturer.id})")
            else:
                # Tạo slug hợp lệ từ manufacturer_name
                manufacturer_slug = re.sub(r'[^a-z0-9-]', '-', manufacturer_name.lower()).strip('-')
                
                # Nếu không có manufacturer, tạo mới
                manufacturer = nb.dcim.manufacturers.create(
                    name=manufacturer_name,
                    slug=manufacturer_slug  
                )
                print(f"Created new manufacturer: {manufacturer.name} (ID: {manufacturer.id})")

            # Tạo slug hợp lệ cho device type
            device_type_slug = re.sub(r'[^a-z0-9-]', '-', device_type.lower()).strip('-')

            # Thêm device type mới
            new_device_type = nb.dcim.device_types.create({
                'model': device_type,
                'slug': device_type_slug,
                'manufacturer': manufacturer.id,
                'u_height': u_height,
                'is_full_depth': 'yes',
            })
            print(f"Automatically added: {device_type}")

        except Exception as e:
            print(f"Error while adding {device_type}: {e}, Data: {row.to_dict()}")
            break


