import csv
import getpass
import os
import re
import time
import pandas as pd
import paramiko

def get_switch_list(csv_file):
    with open(csv_file, newline='') as f:
        return [row[0] for row in csv.reader(f) if row]

def get_interface_status_via_shell(host, username, password):
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, username=username, password=password, look_for_keys=False, allow_agent=False, timeout=10)
        shell = client.invoke_shell()
        time.sleep(1)
        shell.recv(10000)  # Clear banner

        shell.send('terminal length 0\n')
        time.sleep(1)
        shell.recv(10000)  # Clear after terminal length 0

        shell.send('show interface status\n')
        time.sleep(2)
        output = ""
        while shell.recv_ready():
            output += shell.recv(65535).decode(errors='ignore')
            time.sleep(0.5)
        client.close()
        return output
    except Exception as e:
        return f"ERROR: {e}"

def parse_interface_status(output):
    columns = ["Port", "Name", "Status", "Vlan", "Duplex", "Speed", "Type"]
    data = []
    lines = output.splitlines()
    header_found = False
    
    # Find the header line to determine column positions
    header_line = None
    header_positions = {}
    
    for i, line in enumerate(lines):
        if re.match(r"^Port\s+Name\s+Status\s+Vlan\s+Duplex\s+Speed\s+Type", line):
            header_line = line
            header_found = True
            
            # Find the starting position of each column in the header
            current_pos = 0
            for col in columns:
                pos = header_line.find(col, current_pos)
                if pos != -1:
                    header_positions[col] = pos
                    current_pos = pos + len(col)
            
            # Skip the header line
            continue
        
        # Skip lines that consist entirely of dash characters
        if re.match(r"^-+$", line):
            continue
            
        # Process data lines if header has been found
        if header_found:
            # Skip empty lines but don't break the loop
            if not line.strip():
                continue
                
            # Break if we encounter a switch prompt
            if re.match(r"^\S+#", line):
                break
            
            # Parse the line based on column positions
            if len(line) >= header_positions.get("Type", 0):
                row_data = []
                
                # Extract Port (first column)
                port = line[:header_positions["Name"]].strip()
                row_data.append(port)
                
                # Extract Name (may be empty)
                name_end = header_positions["Status"]
                name = line[header_positions["Name"]:name_end].strip()
                row_data.append(name)
                
                # Extract Status
                status_end = header_positions["Vlan"]
                status = line[header_positions["Status"]:status_end].strip()
                row_data.append(status)
                
                # Extract Vlan
                vlan_end = header_positions["Duplex"]
                vlan = line[header_positions["Vlan"]:vlan_end].strip()
                row_data.append(vlan)
                
                # Extract Duplex
                duplex_end = header_positions["Speed"]
                duplex = line[header_positions["Duplex"]:duplex_end].strip()
                row_data.append(duplex)
                
                # Extract Speed
                speed_end = header_positions["Type"]
                speed = line[header_positions["Speed"]:speed_end].strip()
                row_data.append(speed)
                
                # Extract Type (last column)
                type_val = line[header_positions["Type"]:].strip()
                row_data.append(type_val)
                
                data.append(row_data)
    
    return pd.DataFrame(data, columns=columns)


def main():
    # Create int_parsed_outputs directory if it doesn't exist
    output_dir = "int_parsed_outputs"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
    
    # Initialize collections to store all switches and their credentials
    all_switches = []
    switch_credentials = {}
    csv_files = []
    
    # Get first CSV file and credentials
    while True:
        csv_file = input("Enter the path to a CSV file with switch names/IPs: ")
        csv_files.append(csv_file)
        
        # Get credentials for this set of switches
        username = input(f"Enter SSH username for switches in {csv_file}: ")
        password = getpass.getpass(f"Enter SSH password for switches in {csv_file}: ")
        
        # Get switch list
        switches = get_switch_list(csv_file)
        print(f"Found {len(switches)} switches in {csv_file}.")
        
        # Store switches and their credentials
        for switch in switches:
            all_switches.append(switch)
            switch_credentials[switch] = (username, password)
        
        # Ask if user has another CSV file with different credentials
        another = input("Do you have another CSV file with switches that use different credentials? (y/n): ").lower()
        if another != 'y':
            break
    
    # Prepare output file name with timestamp to avoid overwriting existing files
    # Use the first CSV file for naming the output files
    base, ext = os.path.splitext(csv_files[0])
    base_filename = os.path.basename(base)  # Get just the filename without path
    current_date = time.strftime("%Y%m%d")
    
    # Generate unique filename with date and sequence number if needed
    excel_base = os.path.join(output_dir, base_filename + f"_show_int_status_parsed_{current_date}")
    
    # Check if files with this date already exist and add sequence number if needed
    seq_num = 1
    excel_file = f"{excel_base}.xlsx"
    
    while os.path.exists(excel_file):
        seq_num += 1
        excel_file = f"{excel_base}_{seq_num}.xlsx"
    
    print(f"Output will be saved to {excel_file}")

    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        # Process each switch with its corresponding credentials
        for switch in all_switches:
            username, password = switch_credentials[switch]
            print(f"Connecting to {switch}...")
            raw_output = get_interface_status_via_shell(switch, username, password)
            if raw_output.startswith("ERROR:"):
                df = pd.DataFrame([[raw_output]], columns=["Error"])
            else:
                df = parse_interface_status(raw_output)
                if df.empty:
                    df = pd.DataFrame([["No data parsed"]], columns=["Info"])
            sheet_name = str(switch)[:31]  # Excel sheet names limited to 31 chars
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Done! Output saved to {excel_file}")

if __name__ == "__main__":
    main()
