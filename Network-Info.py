import netifaces
import tkinter as tk
import tkinter.ttk as ttk
import subprocess
import re
import xlsxwriter                 
from bs4 import BeautifulSoup
import paramiko
import requests
def get_wifi_info():
    wifi_info = []
    output = subprocess.check_output(["netsh", "wlan", "show", "interfaces"])
    for line in output.decode().split("\n"):
        if "SSID" in line:
            ssid = re.search(r"SSID\s+:\s+(.*)", line).group(1)
            wifi_info.append(f"SSID: {ssid}")
        elif "Protocol" in line:
            protocol = re.search(r"Protocol\s+:\s+(.*)", line).group(1)
            wifi_info.append(f"Protocol: {protocol}")
        elif "Authentication" in line:
            auth = re.search(r"Authentication\s+:\s+(.*)", line).group(1)
            wifi_info.append(f"Security Type: {auth}")
        elif "Channel" in line:
            channel = re.search(r"Channel\s+:\s+(\d+)", line).group(1)
            wifi_info.append(f"Channel: {channel}")
        elif "Radio type" in line:
            radio_type = re.search(r"Radio type\s+:\s+(.*)", line).group(1)
            wifi_info.append(f"Network Band: {radio_type}")
        elif "BSSID" in line:
            bssid = re.search(r"BSSID\s+:\s+(.*)", line).group(1)
            wifi_info.append(f"BSSID: {bssid}")
        elif "Signal" in line:
            signal = re.search(r"Signal\s+:\s+(\d+)%", line).group(1)
            wifi_info.append(f"Signal Strength: {signal}%")
        elif "Link speed" in line:
            link_speed = re.search(r"Link speed \(Receive/Transmit\):\s+(\d+)/(\d+)", line).groups()
            wifi_info.append(f"Link Speed (Receive/Transmit): {link_speed[0]}/{link_speed[1]} (Mbps)")
        
    return wifi_info

def disconnect_wifi():
    try:
        subprocess.check_output(["netsh", "wlan", "disconnect"])
        return ["Disconnected from WiFi network"]
    except subprocess.CalledProcessError as e:
        return ["Error: unable to disconnect from WiFi network"]

def get_network_interfaces():
    interfaces = netifaces.interfaces()
    return interfaces

def get_current_connections():
    connections = []
    output = subprocess.check_output(["netstat", "-an"])
    for line in output.decode().split("\n"):
        if "ESTABLISHED" in line:
            connections.append(line)
    return connections

def get_connected_devices(router_ip, password):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(router_ip, username='admin', password=password)
    stdin, stdout, stderr = ssh.exec_command('arp -a')
    devices = []
    for line in stdout:
        devices.append(line.strip())
    ssh.close()
    return devices



def access_website(router_ip, password):
    url = f"http://192.168.8.1{router_ip}"
    response = requests.get(url, auth=('admin', password))
    soup = BeautifulSoup(response.content, 'html.parser')
    devices = []
    for table in soup.find_all('table'):
        for row in table.find_all('tr'):
            cols = row.find_all('td')
            if len(cols) > 0:
                device = cols[1].text.strip()
                devices.append(device)
    return devices

def show_website_devices(router_ip, password, device_list):
    devices = access_website(router_ip, password)
    device_list.delete(0, tk.END)
    for device in devices:
        device_list.insert(tk.END, device)

def refresh_data(wifi_list, interface_list, connection_list, device_list):
    wifi_list.delete(0, tk.END)
    interface_list.delete(0, tk.END)
    connection_list.delete(0, tk.END)
    device_list.delete(0, tk.END)
    for info in get_wifi_info():
        wifi_list.insert(tk.END, info)
    for interface in get_network_interfaces():
        interface_list.insert(tk.END, interface)
    for connection in get_current_connections():
        connection_list.insert(tk.END, connection)
    for device in get_connected_devices():
        device_list.insert(tk.END, device)

def save_to_excel(wifi_info, interface_info, connection_info, device_info):
    try:
        workbook = xlsxwriter.Workbook('network_info.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'WiFi Info')
        for i, info in enumerate(wifi_info):
            worksheet.write(i+1, 0, info)

        worksheet.write('B1', 'Network Interfaces')
        for i, interface in enumerate(interface_info):
            worksheet.write(i+1, 1, interface)

        worksheet.write('C1', 'Current Connections')
        for i, connection in enumerate(connection_info):
            worksheet.write(i+1, 2, connection)

        worksheet.write('D1', 'Connected Devices')
        for i, device in enumerate(device_info):
            worksheet.write(i+1, 3, device)

        workbook.close()
        print("Excel file saved successfully")
    except Exception as e:
        print("Error saving Excel file:", str(e))

def print_data(wifi_info, interface_info, connection_info, device_info):
    print("WiFi Info:")
    for info in wifi_info:
        print(info)
    print("\nNetwork Interfaces:")
    for interface in interface_info:
        print(interface)
    print("\nCurrent Connections:")
    for connection in connection_info:
        print(connection)
    print("\nConnected Devices:")
    for device in device_info:
        print(device)

import tkinter as tk
from tkinter import ttk

def create_gui():
    root = tk.Tk()
    root.title("Network Information")

    # Create a style for the GUI
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TFrame", background="lightgray")
    style.configure("TLabel", font=("Helvetica", 12), background="lightgray")
    style.configure("TButton", font=("Helvetica", 12), background="lightgray")
    style.configure("TListbox", font=("Helvetica", 12), background="lightgray")

    # Create a main frame
    main_frame = ttk.Frame(root, padding="10 10 10 10")
    main_frame.pack(fill="both", expand=True)

    # Create a left frame for WiFi info and network interfaces
    left_frame = ttk.Frame(main_frame, padding="10 10 10 10")
    left_frame.pack(side="left", fill="both", expand=True)

    ttk.Label(left_frame, text="WiFi Info:", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5)
    wifi_list = tk.Listbox(left_frame, font=("Helvetica", 12), width=30, height=10)
    wifi_list.grid(row=1, column=0, padx=10, pady=5)
    for info in get_wifi_info():
        wifi_list.insert(tk.END, info)

    ttk.Label(left_frame, text="Network Interfaces:", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5)
    interface_list = tk.Listbox(left_frame, font=("Helvetica", 12), width=40, height=10)
    interface_list.grid(row=2, column=0, padx=10, pady=5)
    for interface in get_network_interfaces():
        interface_list.insert(tk.END, interface)

    # Create a middle frame for current connections and connected devices
    middle_frame = ttk.Frame(main_frame, padding="10 10 10 10")
    middle_frame.pack(side="left", fill="both", expand=True)

    ttk.Label(middle_frame, text="Current Connections:", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5)
    connection_list = tk.Listbox(middle_frame, font=("Helvetica", 12), width=60, height=10)
    connection_list.grid(row=1, column=0, padx=10, pady=5)
    for connection in get_current_connections():
        connection_list.insert(tk.END, connection)

    ttk.Label(middle_frame, text="Connected Devices:", font=("Helvetica", 12)).grid(row=0, column=1, padx=10, pady=5)
    device_list = tk.Listbox(middle_frame, font=("Helvetica", 12), width=40, height=10)
    device_list.grid(row=1, column=1, padx=10, pady=5)

    # Create a right frame for router info and buttons
    right_frame = ttk.Frame(main_frame, padding="10 10 10 10")
    right_frame.pack(side="left", fill="both", expand=True)

    ttk.Label(right_frame, text="Router IP:", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5)
    router_ip_entry = ttk.Entry(right_frame, font=("Helvetica", 12), width=20)
    router_ip_entry.grid(row=0, column=1, padx=10, pady=5)

    ttk.Label(right_frame, text="Password:", font=("Helvetica", 12)).grid(row=1, column=0, padx=10, pady=5)
    password_entry = ttk.Entry(right_frame, font=("Helvetica", 12), width=20, show="*")
    password_entry.grid(row=1, column=1, padx=10, pady=5)

    disconnect_button = ttk.Button(right_frame, text="Disconnect from WiFi", command=lambda: wifi_list.insert(tk.END, disconnect_wifi()[0]))
    disconnect_button.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

    refresh_button = ttk.Button(right_frame, text="Refresh", command=lambda: refresh_data(wifi_list, interface_list, connection_list, device_list))
    refresh_button.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

    save_button = ttk.Button(right_frame, text="Save to Excel", command=lambda: save_to_excel(wifi_list.get(0, tk.END), interface_list.get(0, tk.END), connection_list.get(0, tk.END), device_list.get(0, tk.END)))
    save_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

    root.mainloop()

create_gui()