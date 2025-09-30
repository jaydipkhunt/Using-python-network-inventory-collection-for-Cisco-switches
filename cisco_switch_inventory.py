from netmiko import ConnectHandler
from openpyxl import Workbook
from datetime import datetime
import re

# Read device IPs
with open("device_ips.txt") as f:
    device_ips = [line.strip() for line in f if line.strip()]

# SSH credentials
username = "your_username"
password = "your_password"

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Switch Inventory"
ws.append([
    "Hostname",
    "IP Address",
    "Switch#",
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

        # --- Extract Software Version (e.g. 16.09.06) ---
        sw_ver_match = re.search(r"Version\s+(\d+\.\d+\.\d+)", version_output)
        software_version = sw_ver_match.group(1) if sw_ver_match else "N/A"

        # Initialize switch list
        switches = []

        # --- Main Switch (Switch 1) ---
        main_model = re.search(r"Model Number\s+:\s+(WS\S+)", version_output)
        main_serial = re.search(r"System Serial Number\s+:\s+(\S+)", version_output)
        main_mac = re.search(r"Base Ethernet MAC Address\s+:\s+([0-9a-fA-F:.]+)", version_output)

        if main_model and main_serial:
            switches.append({
                "switch_number": 1,
                "model": main_model.group(1),
                "serial": main_serial.group(1),
                "mac": main_mac.group(1) if main_mac else "N/A"
            })

        # --- Other Switches ---
        switch_blocks = re.findall(r"(Switch (\d+)\n[-]+\n(.*?))(?=(\nSwitch \d+)|\Z)", version_output, re.DOTALL)

        for block, number, content, _ in switch_blocks:
            model = re.search(r"Model Number\s+:\s+(WS\S+)", content)
            serial = re.search(r"System Serial Number\s+:\s+(\S+)", content)
            mac = re.search(r"Base Ethernet MAC Address\s+:\s+([0-9a-fA-F:.]+)", content)

            switches.append({
                "switch_number": int(number),
                "model": model.group(1) if model else "N/A",
                "serial": serial.group(1) if serial else "N/A",
                "mac": mac.group(1) if mac else "N/A"
            })

        # Total stack count
        stack_count = len(switches)

        # Write to Excel
        for sw in switches:
            ws.append([
                hostname,
                ip,
                sw["switch_number"],
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
        ws.append([f"Connection Failed", ip, "", "", "", "", "", ""])

# Save Excel file
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"switch_inventory_{timestamp}.xlsx"
wb.save(filename)

print(f"\n📄 Inventory saved to: {filename}")
