import csv
import getpass
import os
import re
import time
import pandas as pd
import paramiko
import networkx as nx
from pyvis.network import Network
from collections import defaultdict

def get_switch_list(csv_file):
    """Read switch hostnames from a CSV file."""
    with open(csv_file, newline='') as f:
        return [row[0] for row in csv.reader(f) if row]

def ssh_to_switch(host, username, password):
    """Establish SSH connection to a switch."""
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, username=username, password=password, look_for_keys=False, allow_agent=False, timeout=10)
        return client
    except Exception as e:
        print(f"Error connecting to {host}: {e}")
        return None

def get_cdp_neighbors(client, host):
    """Get CDP neighbor information from a switch."""
    try:
        shell = client.invoke_shell()
        time.sleep(1)
        shell.recv(10000)  # Clear banner

        shell.send('terminal length 0\n')
        time.sleep(1)
        shell.recv(10000)  # Clear after terminal length 0

        shell.send('show cdp neighbor\n')
        time.sleep(2)
        output = ""
        while shell.recv_ready():
            output += shell.recv(65535).decode(errors='ignore')
            time.sleep(0.5)
        
        return output
    except Exception as e:
        return f"ERROR: {e}"

def parse_cdp_output(output, source_switch):
    """Parse the output of 'show cdp neighbor' command."""
    lines = output.splitlines()
    
    # Find the header line
    header_line_idx = -1
    header_line = ""
    for i, line in enumerate(lines):
        # Check for different variations of the Device ID, Local Interface, and Holdtime headers
        if ("Device ID" in line or "Device-ID" in line or "Device Id" in line) and \
           ("Local Intrfce" in line or "Local Interface" in line) and \
           ("Holdtme" in line or "Hldtme" in line or "Hold Time" in line):
            header_line_idx = i
            header_line = line
            break
    
    if header_line_idx == -1:
        return pd.DataFrame()  # No header found
    
    # Get the positions of each column in the header
    # Check for different variations of the Device ID header
    if "Device ID" in header_line:
        device_id_pos = header_line.find("Device ID")
    elif "Device-ID" in header_line:
        device_id_pos = header_line.find("Device-ID")
    elif "Device Id" in header_line:
        device_id_pos = header_line.find("Device Id")
    else:
        device_id_pos = 0  # Default to the beginning of the line if not found
    
    # Check for different variations of the Local Interface header
    if "Local Intrfce" in header_line:
        local_intrfce_pos = header_line.find("Local Intrfce")
    elif "Local Interface" in header_line:
        local_intrfce_pos = header_line.find("Local Interface")
    else:
        local_intrfce_pos = header_line.find("Local")  # Fallback to just "Local" if full header not found
    
    # Check for different variations of the Holdtime header
    if "Holdtme" in header_line:
        holdtme_pos = header_line.find("Holdtme")
    elif "Hldtme" in header_line:
        holdtme_pos = header_line.find("Hldtme")
    elif "Hold Time" in header_line:
        holdtme_pos = header_line.find("Hold Time")
    else:
        holdtme_pos = header_line.find("Hold")  # Fallback to just "Hold" if full header not found
    capability_pos = header_line.find("Capability")
    platform_pos = header_line.find("Platform")
    
    # Check for different variations of the Port ID header
    if "Port ID" in header_line:
        port_id_pos = header_line.find("Port ID")
    elif "Port Id" in header_line:
        port_id_pos = header_line.find("Port Id")
    else:
        port_id_pos = header_line.find("Port")  # Fallback to just "Port" if full header not found
    
    # Process the data lines
    neighbors = []
    current_device_id = None
    i = header_line_idx + 1
    
    while i < len(lines):
        line = lines[i]
        
        # Skip empty lines or lines with switch prompts
        if not line.strip() or re.match(r"^\S+#", line):
            i += 1
            continue
        
        # Skip separator lines
        if re.match(r"^-+$", line.strip()):
            i += 1
            continue
        
        # Check if this is a device ID line (not indented and not containing interface info)
        if not line.startswith(' ') and "Local Intrfce" not in line:
            current_device_id = line.strip()
            i += 1
            continue
        
        # This is a data line (indented, contains interface and other details)
        if current_device_id and line.startswith(' '):
            # Find where the actual data starts (after indentation)
            data_start_pos = len(line) - len(line.lstrip())
            
            # For indented lines, we need to map the fields based on their positions in the header
            # The indentation shifts everything, so we need to calculate the field widths
            
            # Calculate field widths from the header
            local_intrfce_width = holdtme_pos - local_intrfce_pos
            holdtme_width = capability_pos - holdtme_pos
            capability_width = platform_pos - capability_pos
            platform_width = port_id_pos - platform_pos
            
            # Extract each field using the calculated widths
            local_interface = line[data_start_pos:data_start_pos + local_intrfce_width].strip()
            holdtime = line[data_start_pos + local_intrfce_width:data_start_pos + local_intrfce_width + holdtme_width].strip()
            capability = line[data_start_pos + local_intrfce_width + holdtme_width:data_start_pos + local_intrfce_width + holdtme_width + capability_width].strip()
            platform = line[data_start_pos + local_intrfce_width + holdtme_width + capability_width:data_start_pos + local_intrfce_width + holdtme_width + capability_width + platform_width].strip()
            port_id = line[data_start_pos + local_intrfce_width + holdtme_width + capability_width + platform_width:].strip()
            
            
            neighbors.append({
                'source_switch': source_switch,
                'device_id': current_device_id,
                'local_interface': local_interface,
                'holdtime': holdtime,
                'capability': capability,
                'platform': platform,
                'port_id': port_id
            })
        
        # Handle case where device ID and data are on the same line
        elif len(line.strip()) > 0 and not current_device_id:
            # Extract fields based on column positions
            device_id = line[device_id_pos:local_intrfce_pos].strip()
            local_interface = line[local_intrfce_pos:holdtme_pos].strip()
            holdtime = line[holdtme_pos:capability_pos].strip()
            capability = line[capability_pos:platform_pos].strip()
            platform = line[platform_pos:port_id_pos].strip()
            port_id = line[port_id_pos:].strip()
            
            
            neighbors.append({
                'source_switch': source_switch,
                'device_id': device_id,
                'local_interface': local_interface,
                'holdtime': holdtime,
                'capability': capability,
                'platform': platform,
                'port_id': port_id
            })
        
        i += 1
    
    return pd.DataFrame(neighbors)

def normalize_device_name(device_name):
    """Normalize device names by removing serial numbers in parentheses."""
    # Pattern: hostname(SERIAL) -> hostname
    # Example: toc-o29-vault-sw1(FDO...) -> toc-o29-vault-sw1
    if '(' in device_name and ')' in device_name:
        # Extract the part before the parenthesis
        base_name = device_name.split('(')[0].strip()
        # If the base name ends with a domain, keep it as is
        if base_name.endswith('.umm.edu'):
            return base_name
        # Otherwise, try to find a matching device with domain
        domain_name = f"{base_name}.umm.edu"
        return domain_name
    return device_name

def plot_connections(all_neighbors, output_file):
    """Plot the network connections using PyVis for a more interactive and visually appealing graph."""
    # Create a networkx graph
    G = nx.Graph()
    
    # Create a mapping of original device names to normalized names
    device_name_map = {}
    
    # First pass: build the device name mapping
    for _, row in all_neighbors.iterrows():
        source = row['source_switch']
        target = row['device_id']
        
        # Normalize source and target names
        norm_source = normalize_device_name(source)
        norm_target = normalize_device_name(target)
        
        # Add to mapping
        device_name_map[source] = norm_source
        device_name_map[target] = norm_target
    
    # Second pass: find domain versions of devices
    # This helps consolidate devices that appear both with and without domain
    for orig_name, norm_name in list(device_name_map.items()):
        if not norm_name.endswith('.umm.edu'):
            domain_name = f"{norm_name}.umm.edu"
            if domain_name in device_name_map.values():
                device_name_map[orig_name] = domain_name
    
    # Add nodes and edges with normalized names
    for _, row in all_neighbors.iterrows():
        source = row['source_switch']
        target = row['device_id']
        
        # Use normalized names
        norm_source = device_name_map.get(source, source)
        norm_target = device_name_map.get(target, target)
        
        # Add nodes with device type attribute
        if not G.has_node(norm_source):
            G.add_node(norm_source, device_type='switch')
        
        if not G.has_node(norm_target):
            # Determine device type based on capability
            device_type = 'other'
            # Split capability into individual codes and check for 'R' and 'S'
            capability_codes = row['capability'].split()
            if 'R' in capability_codes:
                device_type = 'router'
            elif 'S' in capability_codes:
                device_type = 'switch'
            G.add_node(norm_target, device_type=device_type)
        
        # Add edge with interface information
        G.add_edge(
            norm_source, 
            norm_target, 
            title=f"{norm_source} ({row['local_interface']}) <-> {norm_target} ({row['port_id']})",
            label=f"{row['local_interface']} → {row['port_id']}",
            local_interface=row['local_interface'],
            port_id=row['port_id']
        )
    
    # Create a PyVis network from the networkx graph
    net = Network(height="900px", width="100%", bgcolor="#ffffff", font_color="black")
    
    # Set physics layout options for better visualization with longer edges for readability
    net.barnes_hut(gravity=-80000, central_gravity=0.3, spring_length=350, spring_strength=0.001, damping=0.09)
    
    # Add the networkx graph to the PyVis network
    net.from_nx(G)
    
    # Define node colors and shapes based on device type
    for node in net.nodes:
        # Get the device type from the original networkx graph
        node_id = node['id']
        if node_id in G.nodes and 'device_type' in G.nodes[node_id]:
            device_type = G.nodes[node_id]['device_type']
            if device_type == 'switch':
                node['color'] = '#4da6ff'  # Blue
                node['shape'] = 'dot'
                node['size'] = 25
            elif device_type == 'router':
                node['color'] = '#59b300'  # Green
                node['shape'] = 'diamond'
                node['size'] = 25
            else:
                node['color'] = '#cccccc'  # Gray
                node['shape'] = 'square'
                node['size'] = 20
        
        # Add hover information
        node['title'] = node['id']
    
    # Add hover information to edges
    for edge in net.edges:
        if 'title' not in edge:
            edge['title'] = f"{edge['from']} <-> {edge['to']}"
    
    # Set options for a more appealing visualization
    net.set_options("""
    {
      "nodes": {
        "font": {
          "size": 14,
          "face": "Tahoma"
        },
        "borderWidth": 2,
        "shadow": true
      },
      "edges": {
        "color": {
          "color": "#848484",
          "highlight": "#1E90FF"
        },
        "width": 2,
        "shadow": true,
        "smooth": {
          "type": "dynamic",
          "roundness": 0.5
        },
        "font": {
          "size": 12,
          "face": "Tahoma",
          "background": "white",
          "strokeWidth": 0,
          "align": "middle"
        },
        "length": 350
      },
      "physics": {
        "stabilization": {
          "iterations": 100
        }
      },
      "interaction": {
        "hover": true,
        "navigationButtons": true,
        "keyboard": true
      }
    }
    """)
    
    # Change the file extension to .html
    html_file = os.path.splitext(output_file)[0] + '.html'
    
    # Get all node IDs for the device list
    all_node_ids = list(G.nodes())
    all_node_ids.sort()  # Sort alphabetically
    
    # Create a custom HTML template with a left pane for device list and search
    html_template = """
    <html>
        <head>
            <meta charset="utf-8">
            <script src="lib/bindings/utils.js"></script>
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.2/dist/dist/vis-network.min.css" integrity="sha512-WgxfT5LWjfszlPHXRmBWHkV2eceiWTOBvrKCNbdgDYTHrT2AeLCGbF4sZlZw3UMN3WtL0tGUoIAKsu8mllg/XA==" crossorigin="anonymous" referrerpolicy="no-referrer" />
            <script src="https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.2/dist/vis-network.min.js" integrity="sha512-LnvoEWDFrqGHlHmDD2101OrLcbsfkrzoSpvtSQtxK3RMnRV0eOkhhBN2dXHKRrUU8p2DGRTk35n4O8nWSVe1mQ==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
            
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-eOJMYsd53ii+scO/bJGFsiCZc+5NDVN2yr8+0RDqr0Ql0h+rP48ckxlpbzKgwra6" crossorigin="anonymous" />
            <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta3/dist/js/bootstrap.bundle.min.js" integrity="sha384-JEW9xMcG8R+pH31jmWH6WWP0WintQrMb4s7ZOdauHnUtxwoG2vI5DkLtS3qm9Ekf" crossorigin="anonymous"></script>
            
            <style type="text/css">
                body {
                    margin: 0;
                    padding: 0;
                    overflow: hidden;
                    height: 100vh;
                }
                
                .container-fluid {
                    height: 100vh;
                    padding: 0;
                }
                
                .row {
                    height: 100%;
                    margin: 0;
                }
                
                #sidebar {
                    background-color: #f8f9fa;
                    padding: 15px;
                    border-right: 1px solid #dee2e6;
                    height: 100%;
                    overflow-y: auto;
                }
                
                #mynetwork {
                    width: 100%;
                    height: 100vh;
                    background-color: #ffffff;
                    position: relative;
                }
                
                .device-list {
                    margin-top: 15px;
                    max-height: calc(100vh - 150px);
                    overflow-y: auto;
                }
                
                .device-item {
                    padding: 8px 12px;
                    border-bottom: 1px solid #dee2e6;
                    cursor: pointer;
                }
                
                .device-item:hover {
                    background-color: #e9ecef;
                }
                
                .device-item.switch {
                    border-left: 4px solid #4da6ff;
                }
                
                .device-item.router {
                    border-left: 4px solid #59b300;
                }
                
                .device-item.other {
                    border-left: 4px solid #cccccc;
                }
                
                h4 {
                    margin-bottom: 15px;
                }
                
                .search-container {
                    margin-bottom: 15px;
                }
                
                #device-search {
                    width: 100%;
                    padding: 8px 12px;
                    border: 1px solid #ced4da;
                    border-radius: 4px;
                }
            </style>
        </head>
        <body>
            <div class="container-fluid">
                <div class="row">
                    <!-- Left sidebar for device list and search -->
                    <div class="col-md-3 col-lg-2" id="sidebar">
                        <h4>Network Devices</h4>
                        <div class="search-container">
                            <input type="text" id="device-search" class="form-control" placeholder="Search devices..." />
                        </div>
                        <div class="device-list" id="device-list">
                            <!-- Device list items will be populated here -->
                        </div>
                    </div>
                    
                    <!-- Network visualization -->
                    <div class="col-md-9 col-lg-10 p-0">
                        <div id="mynetwork"></div>
                    </div>
                </div>
            </div>
            
            <script type="text/javascript">
                // Initialize global variables.
                var edges;
                var nodes;
                var allNodes;
                var allEdges;
                var nodeColors;
                var originalNodes;
                var network;
                var container;
                var options, data;
                var filter = {
                    item : '',
                    property : '',
                    value : []
                };
                
                // This method is responsible for drawing the graph, returns the drawn network
                function drawGraph() {
                    var container = document.getElementById('mynetwork');
                    
                    // Parsing and collecting nodes and edges from the python
                    {nodes_and_edges}
                    
                    nodeColors = {};
                    allNodes = nodes.get({ returnType: "Object" });
                    for (nodeId in allNodes) {
                        nodeColors[nodeId] = allNodes[nodeId].color;
                    }
                    allEdges = edges.get({ returnType: "Object" });
                    // adding nodes and edges to the graph
                    data = {nodes: nodes, edges: edges};
                    
                    var options = {options};
                    
                    network = new vis.Network(container, data, options);
                    
                    // Populate the device list
                    populateDeviceList();
                    
                    // Initialize the search functionality
                    initializeSearch();
                    
                    return network;
                }
                
                // Function to populate the device list in the sidebar
                function populateDeviceList() {
                    const deviceList = document.getElementById('device-list');
                    
                    // Clear existing content
                    deviceList.innerHTML = '';
                    
                    // Get all nodes
                    const nodeIds = Object.keys(allNodes);
                    nodeIds.sort(); // Sort alphabetically
                    
                    // Add each node to the list
                    nodeIds.forEach(nodeId => {
                        const node = allNodes[nodeId];
                        
                        // Create list item
                        const deviceItem = document.createElement('div');
                        deviceItem.className = `device-item ${node.device_type || 'other'}`;
                        deviceItem.textContent = nodeId;
                        deviceItem.setAttribute('data-node-id', nodeId);
                        deviceItem.addEventListener('click', () => focusNode(nodeId));
                        deviceList.appendChild(deviceItem);
                    });
                }
                
                // Function to initialize the search functionality
                function initializeSearch() {
                    const deviceSearch = document.getElementById('device-search');
                    
                    // Filter the device list as the user types
                    deviceSearch.addEventListener('input', function() {
                        filterDeviceList(this.value);
                    });
                    
                    // Handle Enter key press
                    deviceSearch.addEventListener('keydown', function(e) {
                        if (e.key === 'Enter') {
                            // Find the first visible device and focus on it
                            const visibleDevices = document.querySelectorAll('.device-list .device-item[style="display: block;"], .device-list .device-item:not([style*="display: none"])');
                            if (visibleDevices.length > 0) {
                                const nodeId = visibleDevices[0].getAttribute('data-node-id');
                                if (nodeId) {
                                    focusNode(nodeId);
                                }
                            }
                            // Prevent form submission
                            e.preventDefault();
                        }
                    });
                }
                
                // Function to filter the device list
                function filterDeviceList(filterText) {
                    filterText = filterText.toLowerCase();
                    const deviceItems = document.querySelectorAll('.device-list .device-item');
                    let visibleCount = 0;
                    
                    deviceItems.forEach(item => {
                        const deviceName = item.textContent.toLowerCase();
                        if (deviceName.includes(filterText)) {
                            item.style.display = 'block';
                            visibleCount++;
                        } else {
                            item.style.display = 'none';
                        }
                    });
                    
                    return visibleCount;
                }
                
                // Function to focus on a specific node
                function focusNode(nodeId) {
                    // Focus the network on the selected node
                    network.focus(nodeId, {
                        scale: 1.0,
                        animation: {
                            duration: 1000,
                            easingFunction: 'easeInOutQuad'
                        }
                    });
                    
                    // Select the node
                    network.selectNodes([nodeId]);
                    
                    // Highlight the node in the list
                    const listItems = document.querySelectorAll('.device-item');
                    listItems.forEach(item => {
                        if (item.getAttribute('data-node-id') === nodeId) {
                            item.style.backgroundColor = '#e9ecef';
                            item.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                        } else {
                            item.style.backgroundColor = '';
                        }
                    });
                }
                
                // Draw the graph
                drawGraph();
            </script>
        </body>
    </html>
    """
    
    # Replace placeholders in the template
    # Use the original PyVis generated HTML to extract the nodes and edges data
    temp_file = "temp_network.html"
    net.save_graph(temp_file)
    
    with open(temp_file, 'r') as f:
        temp_html = f.read()
    
    # Extract the nodes and edges data from the temporary HTML
    nodes_pattern = r"nodes = new vis.DataSet\(\[(.*?)\]\);"
    edges_pattern = r"edges = new vis.DataSet\(\[(.*?)\]\);"
    
    nodes_match = re.search(nodes_pattern, temp_html, re.DOTALL)
    edges_match = re.search(edges_pattern, temp_html, re.DOTALL)
    
    if nodes_match and edges_match:
        nodes_data = nodes_match.group(1)
        edges_data = edges_match.group(1)
        
        # Insert the extracted data into our custom template
        nodes_and_edges_str = f"nodes = new vis.DataSet([{nodes_data}]);\n                    edges = new vis.DataSet([{edges_data}]);"
        html_content = html_template.replace("{nodes_and_edges}", nodes_and_edges_str)
        
        # Extract the options from the temporary HTML
        options_pattern = r"var options = (.*?);"
        options_match = re.search(options_pattern, temp_html, re.DOTALL)
        
        if options_match:
            options_data = options_match.group(1)
            html_content = html_content.replace("{options}", options_data)
        else:
            # Fallback to the original options string if extraction fails
            html_content = html_content.replace("{options}", net.options)
    else:
        # Fallback to the original method if extraction fails
        nodes_and_edges_str = f"nodes = new vis.DataSet({net.nodes});\n                    edges = new vis.DataSet({net.edges});"
        html_content = html_template.replace("{nodes_and_edges}", nodes_and_edges_str)
        html_content = html_content.replace("{options}", str(net.options).replace("'", '"'))
    
    # Clean up the temporary file
    if os.path.exists(temp_file):
        os.remove(temp_file)
    
    # Write the HTML file
    with open(html_file, 'w') as f:
        f.write(html_content)
    
    print(f"Interactive network visualization with device list saved to {html_file}")
    
    return G

def main():
    # Create cdp_outputs directory if it doesn't exist
    output_dir = "cdp_outputs"
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
    
    # Prepare output file names with date to avoid overwriting existing files
    # Use the first CSV file for naming the output files
    base, ext = os.path.splitext(csv_files[0])
    base_filename = os.path.basename(base)  # Get just the filename without path
    current_date = time.strftime("%Y%m%d")
    
    # Generate unique filenames with date and sequence number if needed
    excel_base = os.path.join(output_dir, base_filename + f"_cdp_neighbors_{current_date}")
    plot_base = os.path.join(output_dir, base_filename + f"_cdp_network_plot_{current_date}")
    
    # Check if files with this date already exist and add sequence number if needed
    seq_num = 1
    excel_file = f"{excel_base}.xlsx"
    plot_file = f"{plot_base}.html"
    
    while os.path.exists(excel_file) or os.path.exists(plot_file):
        seq_num += 1
        excel_file = f"{excel_base}_{seq_num}.xlsx"
        plot_file = f"{plot_base}_{seq_num}.html"
    
    print(f"Output will be saved to {excel_file} and {plot_file}")
    
    # Create Excel writer
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        all_neighbors = pd.DataFrame()
        
        # Process each switch with its corresponding credentials
        for switch in all_switches:
            username, password = switch_credentials[switch]
            print(f"Connecting to {switch}...")
            client = ssh_to_switch(switch, username, password)
            
            if client:
                raw_output = get_cdp_neighbors(client, switch)
                client.close()
                
                if raw_output.startswith("ERROR:"):
                    df = pd.DataFrame([[raw_output]], columns=["Error"])
                else:
                    df = parse_cdp_output(raw_output, switch)
                    if df.empty:
                        df = pd.DataFrame([["No CDP neighbors found"]], columns=["Info"])
                    else:
                        # Add to the combined dataframe
                        all_neighbors = pd.concat([all_neighbors, df], ignore_index=True)
                
                # Save to Excel
                sheet_name = str(switch)[:31]  # Excel sheet names limited to 31 chars
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Processed {switch}: Found {len(df)} CDP neighbors")
            else:
                print(f"Skipping {switch} due to connection error")
        
        # Create a summary sheet with all connections
        if not all_neighbors.empty:
            all_neighbors.to_excel(writer, sheet_name="All_Connections", index=False)
            print(f"Found {len(all_neighbors)} total connections across all switches")
            
            # Plot the connections
            print("Generating network plot...")
            G = plot_connections(all_neighbors, plot_file)
            print(f"Network plot saved to {plot_file}")
        else:
            print("No CDP neighbors found across all switches")
    
    print(f"Done! Output saved to {excel_file}")

if __name__ == "__main__":
    main()
