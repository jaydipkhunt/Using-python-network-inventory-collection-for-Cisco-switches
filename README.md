What This Script Does

This Python script automates the network inventory collection from Cisco switches (especially stackable models like 3850, 3750, etc.) via SSH.

✅ Features

Connects via SSH to each Cisco switch IP from a list (device_ips.txt)

Runs the show version command

Extracts the following information:

✅ Hostname

✅ IP Address

✅ Switch number (for each member in the stack)

✅ Model Number (e.g., WS-C3850-48P)

✅ Serial Number (unique per switch)

✅ MAC Address

✅ Software Version (e.g., 16.09.06)

✅ Stack Count (total number of switches in the stack)

Exports all data to a well-formatted Excel file (.xlsx) with timestamp

Each switch in the stack is logged as a separate row

Handles switches with:

Only 1 unit (standalone)

Multiple members in a stack (up to 9–12 typical)

📦 Output Example
Hostname	IP Address	Switch#	Model Number	Serial Number	MAC Address	Software Version	Stack Count
Core01	10.0.0.1	1	WS-C3850-48P	FCW1234ABC	28:52:61:0e:55:00	16.09.06	4
Core01	10.0.0.1	2	WS-C3850-48P	FOC5678DEF	c4:14:3c:b3:2c:00	16.09.06	4
Core01	10.0.0.1	3	WS-C3850-48P	FOC9012GHI	c4:14:3c:b3:35:80	16.09.06	4
Core01	10.0.0.1	4	WS-C3850-48P	FOC3456JKL	50:1c:bf:9c:7a:00	16.09.06	4
🔒 Security Note

Passwords are stored in plaintext for simplicity — for production use, consider:

Using environment variables

Integrating with secure vaults (e.g., HashiCorp Vault, AWS Secrets Manager)

Prompting for credentials at runtime

🛠️ Script Technologies

Python 3.x

Netmiko
 — for SSH to Cisco switches

OpenPyXL
 — to generate Excel files
