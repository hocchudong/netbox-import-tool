import os
import pandas as pd
import pynetbox
import requests
import re
import urllib3
import warnings
from urllib3.exceptions import InsecureRequestWarning
from openpyxl import load_workbook
import config
from datetime import datetime


filepath = config.filepath
NetBox_URL = config.NetBox_URL
NetBox_Token = config.NetBox_Token
sitename = config.sitename
sheetname = config.sheetname

# Class lưu thông tin độ cao của device type
class DeviceHight:
    def __init__(self, name, height):
        self.name = name
        self.height = height
    
    def __repr__(self):
        return f"DeviceHight(name='{self.name}', height={self.height})"
    
def file_check(input_file):
    if os.path.exists(input_file):
        # read file excel
        workbook = load_workbook(input_file)
        global sheet
        sheet = workbook.active
        global df
        columns = [cell.value for cell in sheet[1]]  
        # new code
        df = pd.read_excel(input_file, sheet_name=sheetname)

        # old code

        #data = [[cell.value for cell in row] for row in sheet.iter_rows(min_row=2)]
        # global df
        # df = pd.DataFrame(data, columns=columns)
        # df = df.dropna(subset=['Name'], how='all')
        # required_columns = ['Role','Type','Serial Number']
        # for index, row in df.iterrows():
        #     if pd.notna(row['Name']):
        #         missing_columns = [col for col in required_columns if pd.isnull(row[col])]
        #         if missing_columns:
        #             print(f"Error: Row {index + 2} is missing values in columns: {missing_columns}")
        #             exit()
        # print("File Check complete!")
    else:
        print(f"File '{input_file}' doesn't exist!")
        exit()

def netbox_connection_check(netboxurl, netboxtoken):
    try:
        warnings.simplefilter("ignore", InsecureRequestWarning)  
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(
            netboxurl,
            headers={"Authorization": f"Token {netboxtoken}"},
            timeout=20,
            verify=False  
        )
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        if response.status_code == 200:
            global nb
            nb = pynetbox.api(netboxurl, token=netboxtoken)
            nb.http_session.verify = False  
            print("Connection Check complete!")
        else:
            print(f"Connection Error: {response.status_code} - {response.reason}")
            return None
    except requests.exceptions.SSLError as e:
        print(f"SSL Error: Can't verify SSL certificate. More: {e}")
    except requests.exceptions.ConnectionError as e:
        print(f"Connection Error: Unable to reach NetBox. More: {e}")
    except requests.exceptions.Timeout as e:
        print(f"Timeout Error: NetBox did not respond in time. More: {e}")
    except requests.exceptions.RequestException as e:
        print(f"Error: An unknown error occurred. More: {e}")
    return None

def handle_duplicate_names(df, name_col, serial_col):
    name_counts = df[name_col].value_counts()
    duplicates = name_counts[name_counts > 1].index  # Lấy các giá trị Name bị trùng

    for name in duplicates:
        duplicate_rows = df[df[name_col] == name].index
        for i, row in enumerate(duplicate_rows, start=1):
            #rack_value = df.at[row, rack_col]
            serial_value = df.at[row, serial_col]
            df.at[row, name_col] = f"{name}_{serial_value}"
    return df

def excute_merge_data(df, sheet):
    try:
        df['u_height'] = 1
        for merged_cells in sheet.merged_cells.ranges:
            min_row, min_col, max_row, max_col = merged_cells.min_row, merged_cells.min_col, merged_cells.max_row, merged_cells.max_col
            if min_col == 3 and max_col == 3:
                merge_height = max_row - min_row + 1
                merged_rows = list(range(min_row - 2, max_row - 1))  
                df.loc[merged_rows, 'u_height'] = merge_height
                if merge_height == 2:
                    df.loc[merged_rows, 'Position'] = df.loc[merged_rows, 'Position'] - 1
                elif merge_height == 3:
                    df.loc[merged_rows, 'Position'] = df.loc[merged_rows, 'Position'] - 2
                elif merge_height == 6:
                    df.loc[merged_rows, 'Position'] = df.loc[merged_rows, 'Position'] - 5
        return df
    except Exception as e:
        print(f"Error while excuting merge_data: {e}")
        exit()   
    
def site_check(site_name):
    site_record = nb.dcim.sites.get(name=site_name)
    if site_record:
        print("Site check complete!")
    else:
        print(f"Site '{site_name}' does not exist in NetBox!") 
        print("Do you want to auto add new Site?")
        choice = input("yes/no: ").strip().lower()
        if choice == 'yes':
            new_site = nb.dcim.sites.create(
                name=site_name,
                slug=site_name.lower().replace(" ", "-"),
                status=config.status,
                description='Create by Auto_Import_Tool',
            )
            print(f"Successfully created Site: {site_name}")
        else:
            print("Please check the Site again before using this Tool!")

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

def device_role_check():
    # device_role_names = df['Role'].dropna().drop_duplicates().apply(get_role).tolist()
    device_role_names = df["Role"].tolist() # chuyển pandas thành list 
    device_role_names = list(set(device_role_names)) # lọc trùng
    device_role_names = list(filter(lambda x: x == x, device_role_names)) # loại bỏ phần tử nan trong chuỗi

    all_roles_exist = True # tất cả đã tồn tại chưa
    missing_roles = [] # danh sách các roles chưa có 

    # tìm kiếm các role bị thiếu
    for device_roles in device_role_names:
        dvr = nb.dcim.device_roles.filter(name=device_roles)
        if not dvr:
            print(f"Device Role: {device_roles} does not exist in NetBox, please check again!")
            all_roles_exist = False  
            missing_roles.append(device_roles)
    
    # kiểm tra tất cả role đã tồn tại hay chưa
    if all_roles_exist:
        print("Device Role check complete: All roles exist in NetBox!")
    else:
        print("Do you want to auto add Role?")
        choice = input("yes/no: ").strip().lower()

        if choice == 'yes':
            for role in missing_roles:
                try:
                    new_role = nb.dcim.device_roles.create(
                        name=role,
                        slug=role.lower().replace(" ", "-"),
                        color="9e9e9e",
                        description='Create by Auto_Import_Tool',  
                    )
                    print(f"Successfully created Device Role: {new_role['name']}")
                except Exception as e:
                    print(f"Failed to create Device Role {role}: {e}")
        else:
            print("No roles were added. Please update NetBox manually if needed.")
            exit()
        
def rack_check():
    rack_names = df['Rack'].drop_duplicates().tolist()
    missing_racks = []
    for rack_name in rack_names:
        record = nb.dcim.racks.get(name=rack_name)
        if record:
            continue
        else:
            print(f"Rack '{rack_name}' does not exist in NetBox!")
            missing_racks.append(rack_name)
    
    if missing_racks:
        print("Do you want to auto add new Rack?")
        choice = input("yes/no: ").strip().lower()
        if choice == 'yes':
            for rack in missing_racks:
                try:
                    site = nb.dcim.sites.get(name=config.sitename)
                    if not site:
                        print(f"Site '{config.sitename}' does not exist in NetBox. Please create it first.")
                        break
                    new_rack = nb.dcim.racks.create(
                        site=site.id,
                        name=rack,
                        status='active',
                        width=config.width,
                        u_height=config.u_height,
                        description='Create by Auto_Import_Tool',
                    )
                    print(f"Successfully created Rack: {new_rack['name']}")
                except Exception as e:
                    print(f"Failed to create Rack {rack}: {e}")
        else:
            print("No racks were added. Please update NetBox manually if needed.")
    else:
        print("Rack Check complete: All racks exist in NetBox!")

def manufacturer_check():
    manufacturer_names = df["Manufacturer"].dropna().tolist()
    manufacturer_names = list(set(manufacturer_names))# lọc trùng

    all_manufacturers_exist = True  
    missing_manufacturers = []
    for manufacturer in manufacturer_names:
        dvr = nb.dcim.manufacturers.filter(name=manufacturer)
        if not dvr:
            print(f"Manufacturer: {manufacturer} does not exist in NetBox, please check again!")
            all_manufacturers_exist = False  
            missing_manufacturers.append(manufacturer)

    if all_manufacturers_exist:
        print("Device manufacturer check complete: All manufacturers exist in NetBox!")
    else:
        print("Do you want to auto add manufacturer?")
        choice = input("yes/no: ").strip().lower()

        if choice == 'yes':
            for manufacturer in missing_manufacturers:
                try:
                    new_manufacturer = nb.dcim.manufacturers.create(
                        name=manufacturer,
                        slug=manufacturer.lower().replace(" ", "-"),
                        description='Create by Auto_Import_Tool',  
                    )
                    print(f"Successfully created Device Role: {new_manufacturer['name']}")
                except Exception as e:
                    print(f"Failed to create Device Role {manufacturer}: {e}")
        else:
            print("No manufacturers were added. Please update NetBox manually if needed.")
            exit()

def custom_feild_check():
    custom_feild_names = ["device_owner", "contract_number" ,"year_of_investment"]
    all_custom_feild_exist = True  
    missing_custom_feild = []
    for custom_feild in custom_feild_names:
        dvr = nb.extras.custom_fields.filter(name=custom_feild)
        if not dvr:
            print(f"custom_feild: {custom_feild} does not exist in NetBox, please check again!")
            all_custom_feild_exist = False  
            missing_custom_feild.append(custom_feild)

    if all_custom_feild_exist:
        print("Device custom_feild check complete: All custom_feild exist in NetBox!")
    else:
        print("Do you want to auto add custom_feild?")
        choice = input("yes/no: ").strip().lower()

        if choice == 'yes':
            for custom_feild in missing_custom_feild:
                try:
                    if custom_feild == "year_of_investment":
                        type_custom_feild = "datetime"
                        new_custom_feild = nb.extras.custom_fields.create(
                            name = custom_feild,
                            type = type_custom_feild,
                            object_types = ["dcim.device"],
                            search_weight = 1000,
                            weight = 100,
                            filter_logic = "loose",
                            ui_visible = "always",
                            ui_editable = "yes",
                            description = 'Create by Auto_Import_Tool',
                        )
                    else:
                        type_custom_feild = "text"
                        new_custom_feild = nb.extras.custom_fields.create({
                            "name" : custom_feild,
                            "type" : type_custom_feild,
                            "object_types" : ["dcim.device"],
                            "search_weight" : 1000,
                            "weight" : 100,
                            "filter_logic" : "loose",
                            "ui_visible" : "always",
                            "ui_editable" : "yes",
                            "description" : 'Create by Auto_Import_Tool',
                        })

                    print(f"Successfully created custom feild : {new_custom_feild['name']}")
                except Exception as e:
                    print(f"Failed to create custom feild {custom_feild}: {e}")
        else:
            print("No custom_feild were added. Please update NetBox manually if needed.")
            exit()

#Hàm tạo ra mảng chứa trường name và u_height của device_type
def device_type_height():
    global device_heights
    device_heights = []
    # Tìm u_height của device type 
    print("Merged cell ranges in column H spanning rows:")
    for merged_range in sheet.merged_cells.ranges:
        # Lấy giá trị ở cột  H
        if merged_range.min_col == 8 and merged_range.max_col != merged_range.max_row:
            # lấy value của ô 
            top_left_cell = sheet.cell(merged_range.min_row, merged_range.min_col)
            device_type_name = top_left_cell.value

            # Tính các row đã merg của ô
            rows_spanned = merged_range.max_row - merged_range.min_row + 1

            exists = any(device.name == device_type_name for device in device_heights)
            # kiểm tra giá trị đã có trong list chưa
            if not exists:
                new_device_height =  DeviceHight(device_type_name,rows_spanned)
                device_heights.append(new_device_height)     

# Hàm auto add device type vào netbox
def device_types_check():
    device_types_in_file = df['Type'].dropna().tolist() 
    device_types_in_file = list(set(device_types_in_file)) # lọc trùng
    device_type_not_in_netbox = []

    # Tìm phần tử chưa có trong netbox
    for device_type in device_types_in_file:
        search_result = nb.dcim.device_types.filter(model=device_type)
        if not search_result:
            device_type_not_in_netbox.append(device_type)

    if device_type_not_in_netbox:
        print("Device Types not in NetBox:")
        print(device_type_not_in_netbox)
        print("\nDo you want to add Device Type automatically?")
        choice = input("Enter your choice? (yes/no): ")

        if choice == 'no':
            print("\nPlease Add Device Types manually!")
            exit()
        elif choice == "yes":
            print("You chose to Add Device Types Automatically with sample information")
            print("Trying to Add Automatically...")
            for device_type in device_type_not_in_netbox:
                try:
                    # Tìm manufature của device type trong danh sách 
                    all_manufacturer_device_type = df[["Manufacturer","Type"]].dropna()
                    manufacturer_device_type = all_manufacturer_device_type.query(f"Type == '{device_type}'")
                    manufacturer_name = manufacturer_device_type.iloc[0]['Manufacturer']
                    manufacturer = nb.dcim.manufacturers.get(name = manufacturer_name)

                    # Lấy chiều cao trong list đã tạo
                    device_height = 1 # mặc định height bằng 1 
                    
                    #Tìm kiếm trong danh sách thiết bị có height > 1
                    for device in device_heights:
                        if device.name == device_type:
                            device_height = device.height
                            break
                    
                    # create device type
                    device_type_slug = re.sub(r'[^a-z0-9-]', '-', device_type.lower()).strip('-')
                    new_device_type = nb.dcim.device_types.create({
                        'model': device_type,
                        'slug': device_type_slug,
                        'manufacturer': manufacturer.id,
                        'u_height': device_height,
                        'is_full_depth': 'yes',
                    })

                    print(f"Automatically added : {new_device_type}")
                    
                    # old code

                    # matching_rows = df[df['Type'] == device_type]
                    # if matching_rows.empty:
                    #     print(f"Device type {device_type} not found in Excel. Skipping...")
                    #     continue
                    # row = matching_rows.iloc[0]
                    # u_height = row['u_height']
                    # u_height = int(u_height)
                    # manufacturer_name = row['Manufacturer'].strip()
                    # manufacturer_records = nb.dcim.manufacturers.filter(name=manufacturer_name)
                    # manufacturer = None
                    # for record in manufacturer_records:
                    #     manufacturer = record
                    #     break  
                    # if manufacturer:
                    #     print(f"Using existing manufacturer: {manufacturer.name} (ID: {manufacturer.id})")
                    # else:
                    #     manufacturer_slug = re.sub(r'[^a-z0-9-]', '-', manufacturer_name.lower()).strip('-')
                    #     manufacturer = nb.dcim.manufacturers.create(
                    #         name=manufacturer_name,
                    #         slug=manufacturer_slug  
                    #     )
                    #     print(f"Created new manufacturer: {manufacturer.name} (ID: {manufacturer.id})")
                    # device_type_slug = re.sub(r'[^a-z0-9-]', '-', device_type.lower()).strip('-')
                    # new_device_type = nb.dcim.device_types.create({
                    #     'model': device_type,
                    #     'slug': device_type_slug,
                    #     'manufacturer': manufacturer.id,
                    #     'u_height': u_height,
                    #     'is_full_depth': 'yes',
                    # })

                    
                except Exception as e:
                    print(f"Error while adding {device_type}: {e}")
                    exit()
    else:
        print("Device Types check complete!")
    
def get_device_types_ids(device_types_names):
    try:
        device_types_names = device_types_names.strip()
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
        device_role_name = device_role_name.strip()
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
        rack_name = rack_name.strip()
        rack = nb.dcim.racks.get(name=rack_name)
        return rack.id
    except Exception as e:
        print(f"Error fetching rack '{rack_name}': {e}")
        return None

# Hàm auto add device vào netbox
def import_device_to_NetBox():
    # list device need add 
    device_names = df[["Rack", "Position", "Manufacturer", "Name", "Role", "Device Owner",  "Contract number","Type", "Serial Number","Year of Investment", "Description"]]
    device_names= device_names[device_names['Name'].notna()] # lọc tất cả hàng có trường Name là nan

    number_of_device_in_file = 0
    number_of_device_has_been_added = 0
    for index, row in device_names.iterrows():
        number_of_device_in_file+=1
        # kiểm tra xem device name đã được add trên netbox chưa
        device_name = row['Name']
        device_name = device_name.strip()
        exist_device = nb.dcim.devices.get(name=device_name)
        if exist_device:
            device_name = device_name + f"_{row['Serial Number']}"

        # Thêm mới device đến netbox
        rack_id = get_rack_id(row['Rack'])
        site_id = get_site_id(site_name=sitename)
        device_types_id = get_device_types_ids(row['Type'])
        device_roles_id = get_device_roles_ids(row['Role'])

        # Xử lý trường position
        device_position = row['Position']
        #Tìm kiếm trong danh sách thiết bị có height > 1
        for device in device_heights:
            if device.name == row['Type']:
                device_position = device_position - device.height + 1
                break

        try:
            #Kiểm tra nan cho các param 
            device_description = ""
            device_owner = ""
            device_contract_number = ""
            device_year_of_investment = ""

            if not pd.isna(row['Description']):
                device_description = row['Description']

            if not pd.isna(row['Device Owner']):
                device_owner = row['Device Owner']

            if not pd.isna(row['Contract number']):
                contract_number = row['Contract number']

            if not pd.isna(row['Year of Investment']):
                device_year_of_investment = row['Year of Investment']
                # convert string to data time format YYYY-MM-DD HH:MM:SS
                date_object = datetime.strptime(device_year_of_investment, "%m/%d/%Y")
                formatted_date = date_object.strftime("%Y-%m-%d %H:%M:%S")
                device_year_of_investment = formatted_date

            new_device = nb.dcim.devices.create(
                {
                    "name": device_name,
                    "device_type": device_types_id,
                    "role": device_roles_id,
                    "site": site_id,
                    "serial": row['Serial Number'],
                    "rack": rack_id,
                    "face": "front",
                    "position": device_position,
                    "status": config.status,  
                    "description": device_description,
                    "custom_fields": {
                        "device_owner": device_owner,
                        "contract_number": contract_number,
                        "year_of_investment": device_year_of_investment,
                    }
                }
            )
            
            number_of_device_has_been_added+=1
            print(f"Successfully created device: {device_name}")
        except Exception as e:
            print(f"Error creating device '{device_name}': {e}")
            
    if number_of_device_has_been_added > 0:
        print(f"{number_of_device_has_been_added}/{number_of_device_in_file} device has been added to NetBox!" )

def main():
    try:
        print("Step 1: Checking input file...")
        file_check(filepath)
        
        print("Step 2: Checking NetBox connection...")
        netbox_connection_check(NetBox_URL, NetBox_Token)

        print("Step 3: check site name")
        site_check(site_name=sitename)

        print("Step 4: check rack")
        rack_check()

        print("Step 5: check device role")
        device_role_check()
        
        print("Step 6: check manufacturer")
        manufacturer_check()

        print("Step 7: check custom feild")
        custom_feild_check()
        
        print("Step 8: Check height device type")
        device_type_height()

        print("Step 9: check device type")
        device_types_check()

        print("Step 10: Importing Devices into NetBox...")
        import_device_to_NetBox()

        # print("Step 4: Excuting Merge data...")
        # excute_merge_data(df=df,sheet=sheet)

        print("Process completed successfully!")

    except Exception as e:
        print(f"Error during execution: {e}")

if __name__ == "__main__":
    main()

