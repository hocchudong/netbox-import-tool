import os
import pandas as pd
import pynetbox
import requests
import re
import urllib3
import config

filepath = config.filepath
sheetname = config.sheetname
NetBox_URL = config.NetBox_URL
NetBox_Token = config.NetBox_Token
sitename = config.sitename


def file_check():
    if os.path.exists(filepath):
        global df
        df = pd.read_excel(filepath, sheet_name=sheetname)
        print("File Check complete!")
    else:
        print(f"File '{filepath}' doesn't exist!")
        exit()
        
def netbox_connection_check():
    try:
        response = requests.get(
            NetBox_URL,
            headers={"Authorization": f"Token {NetBox_Token}"},
            timeout=20
        )
        if response.status_code == 200:
            global nb
            nb = pynetbox.api(NetBox_URL, token=NetBox_Token)
            nb.http_session.verify = False                                      
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            print("Connection Check complete!")
        else:
            print(f"Connection Error: {response.status_code} - {response.reason}")
    except requests.exceptions.RequestException as e:
        print(f"Error: Can't connect NetBox. More: {e}")
        
def device_types_check():
    # Lấy ra danh sách các device types từ file xlsx
    device_types_in_file = df['Device Type'].dropna().tolist() 
    
    # Khởi tạo mảng chứa
    device_types_not_in_netbox = []
    # Kiểm tra xem mỗi device type đã có trong NetBox chưa
    for device_type in device_types_in_file:
        search_result = nb.dcim.device_types.filter(model=device_type)
        if not search_result:
            device_types_not_in_netbox.append(device_type)
    # In ra danh sách device types chưa có trên NetBox
    if device_types_not_in_netbox:
        print("Device Types not in NetBox:")
        print(device_types_not_in_netbox)

        # Chọn cách xử lý cho các device types chưa có trên NetBox
        print("\nWhat do you want to do with the device_types not in NetBox")
        print("1. Add manual and relaunch later!")
        print("2. Automatic add with sample value")

        choice = input("Enter your choice? (1 or 2): ")

        if choice == '1':
            print("\nPlease Add Device Types manually!")
            exit()
        elif choice == "2":
            print("You chose to Add Device Automatically with sample information")
            print("Trying to Add Automatically...")

            # Thêm device types chưa có trong NetBox
            for device_type in device_types_not_in_netbox:
                try:
                    # Lấy thông tin từ file xlsx
                    matching_rows = df[df['Device Type'] == device_type]
                    if matching_rows.empty:
                        print(f"Device type {device_type} not found in Excel. Skipping...")
                        continue
                    
                    row = matching_rows.iloc[0]
                    manufacturer_name = row['Manufacturer'].strip()

                    # Kiểm tra manufacturer tồn tại trên NetBox
                    manufacturer_records = nb.dcim.manufacturers.filter(name=manufacturer_name)
                    manufacturer = None
                    for record in manufacturer_records:
                        manufacturer = record
                        break  # Lấy manufacturer đầu tiên tìm thấy
                    
                    if manufacturer:
                        print(f"Using existing manufacturer: {manufacturer.name} (ID: {manufacturer.id})")
                    else:
                        # Tạo slug hợp lệ từ manufacturer_name
                        manufacturer_slug = re.sub(r'[^a-z0-9-]', '-', manufacturer_name.lower()).strip('-')

                        # Tạo manufacturer mới nếu không tồn tại
                        manufacturer = nb.dcim.manufacturers.create(
                            name=manufacturer_name,
                            slug=manufacturer_slug  
                        )
                        print(f"Created new manufacturer: {manufacturer.name} (ID: {manufacturer.id})")

                    # Tạo slug hợp lệ cho device type
                    device_type_slug = re.sub(r'[^a-z0-9-]', '-', device_type.lower()).strip('-')

                    # Thêm device type mới vào NetBox
                    new_device_type = nb.dcim.device_types.create({
                        'model': device_type,
                        'slug': device_type_slug,
                        'manufacturer': manufacturer.id,
                        'u_height': 1,
                        'is_full_depth': 'yes',
                    })

                    print(f"Automatically added: {device_type}")

                except Exception as e:
                    print(f"Error while adding {device_type}: {e}, Data: {row.to_dict()}")
                    exit()
    else:
        print("Device Types check complete!")
        
def site_check(site_name):
    site_record = nb.dcim.sites.get(name=site_name)
    if site_record:
        print("Site check complete!")
    else:
        raise("Site doesn't Exist!")
        
'''
def get_role(role_value):
        if isinstance(role_value, str):  
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
        return None  
'''

def device_role_check():
    # Xử lý danh sách device roles từ file Excel
    device_role_names = df['Role'].dropna().tolist()
    
    all_roles_exist = True  # Biến để kiểm tra xem tất cả các role có tồn tại hay không

    for device_roles in device_role_names:
        # Tìm kiếm device role trong NetBox
        dvr = nb.dcim.device_roles.filter(name=device_roles)
        if not dvr:
            print(f"Device Role: {device_roles} does not exist in NetBox, please check again!")
            all_roles_exist = False  # Nếu không tồn tại thì đặt all_roles_exist = False
    
    if all_roles_exist:
        print("Device Role check complete: All roles exist in NetBox!")
    else:
        print("Device Role check complete with missing roles.")
        
def rack_check():
    rack_name = df['Rack'].dropna().drop_duplicates()
    record = nb.dcim.racks.get(name=rack_name)
    if record:
        print("Rack Check complete!")
    else:
        raise("Error while finding rack in NetBox")
    
def get_device_types_ids(device_types_names):
    try:
        device_types = nb.dcim.device_types.filter(name=device_types_names)
        if device_types:
            for device_type in device_types:
                if device_type.model == device_types_names:
                    return device_type.id
        else:
            print(f"Device type '{device_types_names}' not found in NetBox.")
            return None
    except Exception as e:
        print(f"Error fetching device type '{device_types_names}': {e}")
        return None
    
def get_device_roles_ids(device_role_name):
    try:
        device_roles = nb.dcim.device_roles.filter(name=device_role_name)
        if device_roles:
            for device_role in device_roles:
                if device_role.name == device_role_name:
                    return device_role.id
        else:
            print(f"Device Role '{device_role_name}' not found in NetBox.")
            return None
    except Exception as e:
        print(f"Error fetching device role'{device_role_name}': {e}")
        return None
    
def get_site_id(site_name):
    try:
        site = nb.dcim.sites.get(name=site_name)
        return site.id
    except Exception as e:
        print(f"Error fetching site '{site_name}': {e}")
        return None
    
def get_rack_id(rack_name):
    try:
        rack = nb.dcim.racks.get(name=rack_name)
        return rack.id
    except Exception as e:
        print(f"Error fetching rack '{rack_name}': {e}")
        return None
    
def import_device_to_NetBox():
    device_names = df['Name'].dropna().tolist()
    for device_name in device_names:
        matching_rows = df[df['Name'] == device_name]
        if matching_rows.empty:
            print(f"Device Name {device_name} not found in Excel. Skipping...")
            continue
        
        row = matching_rows.iloc[0]
        device_role = row.get('Role')
        name = row.get('Name')
        rack = row.get('Rack')
        device_types = row.get('Device Type')
        serial_number = row.get('Serial Number', ' ')
        comments = row.get('Comments', ' ')
        contract_number = row.get('Contract Number', ' ')
        year_of_investment = row.get('Year of Investment', ' ')
        # Ép kiểu sang dạng int
        position = int(row.get('U'))
        
        device_types_id = get_device_types_ids(device_types)

        device_roles_id = get_device_roles_ids(device_role)

        rack_id = get_rack_id(rack)
        site_id = get_site_id(sitename)
        
        try:
            new_device = nb.dcim.devices.create(
                {
                    "name": name,
                    "device_type":device_types_id,
                    "role": device_roles_id,
                    "site": site_id,
                    "serial": serial_number,
                    "rack": rack_id,
                    "face": "front",
                    "position": position,
                    "status": "active",  
                    "comments": comments,
                    "contract_number": contract_number,
                    "year_of_investment": year_of_investment,
                }
            )
            print(f"Successfully created device: {name}")
        except Exception as e:
            print(f"Error creating device '{name}': {e}")

def main():
    try:
        # Kiểm tra file đầu vào
        print("Step 1: Checking input file...")
        file_check(filepath, sheetname)
        
        # Kiểm tra kết nối NetBox
        print("Step 2: Checking NetBox connection...")
        netbox_connection_check(NetBox_URL, NetBox_Token)
        
        # Kiểm tra Device Types
        print("Step 3: Checking Device Types...")
        device_types_check()

        # Kiểm tra Device Roles
        print("Step 4: Checking Device Roles...")
        device_role_check()

        # Kiểm tra Rack
        print("Step 5: Checking Rack...")
        rack_check()
        
        # Import Devices vào NetBox
        print("Step 6: Importing Devices into NetBox...")
        import_device_to_NetBox()

        print("Process completed successfully!")

    except Exception as e:
        print(f"Error during execution: {e}")

if __name__ == "__main__":
    main()

