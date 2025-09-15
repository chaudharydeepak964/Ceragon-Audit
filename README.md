# Ceragon-Audit

This project generates an audit report of Ceragon microwave links based on the CM (Configuration Management) file. The report identifies the total number of links that need to be audited, validates configuration consistency, and highlights discrepancies for further action. I use this in my current Organisation VI.

🔧 Repository Structure

The repo contains two main scripts:

SFTP (Extract): Fetches CM files from the remote server.

Transform & Load: Processes the extracted data, validates configurations, and generates the final audit report.

🔐 Note on Security

The code has been modified to replace server IPs and file paths due to security policy restrictions.
👉 You can update these values with your own server details before execution.