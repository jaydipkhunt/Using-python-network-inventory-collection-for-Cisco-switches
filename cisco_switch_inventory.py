from netmiko import ConnectHandler
from openpyxl import Workbook
from datetime import datetime
from getpass import getpass
import re

# Prompt credentials securely
username = input("Enter username: ")
password = getpass("Enter password: ")

# Read device IPs
with open("device_ips.txt") as f:
    device_ips = [line.strip() for line in f if line.strip()]

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Switch Inventory"
ws.append([
    "Hostname",
    "IP Address",
    "Switch#",
    "Role",
    "Model Number",
    "Serial Number",
    "MAC Address",
    "Software Version",
    "Stack Count"
])

for ip in device_ips:
    print(f"\n🔄 Connecting to {ip}...")

    device = {
        "device_type": "cisco_ios",
        "host": ip,
        "username": username,
        "password": password,
        "fast_cli": True,
    }

    try:
        connection = ConnectHandler(**device)

        # Get hostname
        hostname = connection.find_prompt().replace("#", "").strip()

        # Run show version
        version_output = connection.send_command("show version")

        # Extract software version
        sw_ver_match = re.search(r"Version\s+([\d.()A-Za-z]+)", version_output)
        software_version = sw_ver_match.group(1) if sw_ver_match else "N/A"

        switches = []

        # ---- MASTER SWITCH (Top section) ----
        master_model = re.search(r"Model Number\s+:\s+(\S+)", version_output)
        master_serial = re.search(r"System Serial Number\s+:\s+(\S+)", version_output)
        master_mac = re.search(r"Base Ethernet MAC Address\s+:\s+([0-9a-fA-F:.]+)", version_output)

        if master_model and master_serial:
            switches.append({
                "switch_number": 1,
                "role": "Master",
                "model": master_model.group(1),
                "serial": master_serial.group(1),
                "mac": master_mac.group(1) if master_mac else "N/A"
            })

        # ---- MEMBER SWITCHES (Switch 2, 3, etc.) ----
        switch_blocks = re.findall(r"(Switch \d+\n[-]+\n.*?)(?=(\nSwitch \d+)|\Z)", version_output, re.DOTALL)

        for block, _ in switch_blocks:
            sw_num_match = re.search(r"Switch\s+(\d+)", block)
            model_match = re.search(r"Model Number\s+:\s+(\S+)", block)
            serial_match = re.search(r"System Serial Number\s+:\s+(\S+)", block)
            mac_match = re.search(r"Base Ethernet MAC Address\s+:\s+([0-9a-fA-F:.]+)", block)

            switches.append({
                "switch_number": sw_num_match.group(1) if sw_num_match else "N/A",
                "role": "Member",
                "model": model_match.group(1) if model_match else "N/A",
                "serial": serial_match.group(1) if serial_match else "N/A",
                "mac": mac_match.group(1) if mac_match else "N/A"
            })

        # ---- Stack count ----
        stack_count = len(switches)

        # ---- Write to Excel ----
        for sw in switches:
            ws.append([
                hostname,
                ip,
                sw["switch_number"],
                sw["role"],
                sw["model"],
                sw["serial"],
                sw["mac"],
                software_version,
                stack_count
            ])

        connection.disconnect()
        print(f"✅ Done: {hostname} ({stack_count} switches)")

    except Exception as e:
        print(f"❌ Failed to connect to {ip}: {e}")
        ws.append([f"Connection Failed", ip, "", "", "", "", "", "", ""])

# Save Excel file
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"switch_inventory_{timestamp}.xlsx"
wb.save(filename)

print(f"\n📄 Inventory saved to: {filename}")
