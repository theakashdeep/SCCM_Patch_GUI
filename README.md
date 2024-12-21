# SCCM Patching GUI Script

This repository contains a PowerShell script designed to streamline patch management in SCCM (System Center Configuration Manager). The script provides a graphical user interface (GUI) for easier operation and enhanced functionality.

---

## Features
- Validates user access based on Active Directory groups.
- Automates site selection for SCCM configuration.
- Integrates with SCCM PowerShell module for advanced operations.

---

## Prerequisites
1. **Active Directory (AD) Access**:
   - Ensure your user account is part of the required AD groups.
   - Replace the placeholder group names (`AD_Group_1`, `AD_Group_2`) with the appropriate group names for your environment.

2. **SCCM PowerShell Module**:
   - The script depends on the SCCM PowerShell module located at:
     ```
     C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1
     ```
   - Update this path to match the location of the SCCM module on your system.

3. **Domain Configuration**:
   - Replace the placeholder domain name (`mysite.mydomain.com`) with your SCCM server’s actual FQDN (Fully Qualified Domain Name).

---

## Usage Instructions
1. **Clone the Repository**:
   ```
   git clone https://github.com/yourusername/SCCM_Patching_GUI.git
   cd SCCM_Patching_GUI
   ```

2. **Modify the Script**:
   - Open the script file (`SCCM_Patching_GUI.ps1`) in your preferred editor.
   - Replace placeholders with your environment-specific details:
     - AD group names.
     - SCCM module path.
     - Domain name.

3. **Run the Script**:
   - Launch PowerShell with administrative privileges.
   - Execute the script:
     ```
     .\SCCM_Patching_GUI.ps1
     ```

---

## Notes
- Ensure all paths and configuration details in the script are customized for your environment before execution.
- The script is provided "as is" without warranty of any kind. Use at your own risk.

---

## Contributing
Contributions are welcome! If you encounter issues or have suggestions for improvement, please open an issue or submit a pull request.

---

## License
This project is licensed under the GNU General Public License v3.0. See the LICENSE file for details.

---

## Disclaimer
This script is intended for educational and informational purposes only. It is the user’s responsibility to ensure compliance with organizational policies and best practices.

