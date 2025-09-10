🎯 Features

Dual-Function GUI: Single tool for both AD management and Entra ID synchronization
Bulk Operations: Update hundreds of computers in minutes
Extension Attributes 1-15: Full support for all AD extension attributes
Preview Mode: Test changes safely before applying
Graph API Integration: Direct sync to Entra ID via Microsoft Graph
Smart Matching: Matches devices by SID or display name
Export Capabilities: Generate CSV reports of all operations
Auto-Discovery: Automatically detects domain and common OU structures
Offline Support: Works in air-gapped environments with offline modules

📋 Prerequisites
Required Permissions

Active Directory: Read/Write access to computer objects
Entra ID: Device.ReadWrite.All permission in Microsoft Graph
Local Machine: Administrator rights (for module installation)

System Requirements

Windows 10/11 or Windows Server 2016+
PowerShell 7 or later
.NET Framework 4.7.2+
Domain-joined machine

PowerShell Modules
The tool will auto-install these if missing:

ActiveDirectory (RSAT)
Microsoft.Graph.Authentication
Microsoft.Graph.Identity.DirectoryManagement

🚀 Quick Start
1. Download the Tool
powershell# Clone the repository
git clone https://github.com/YOUR-USERNAME/AD-Entra-Extension-Manager.git

# Or download the script directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YOUR-USERNAME/AD-Entra-Extension-Manager/main/AD-EntraExtensionManager.ps1" -OutFile "AD-EntraExtensionManager.ps1"
2. Run the Tool
powershell# Basic usage
.\AD-EntraExtensionManager.ps1

# With proxy
.\AD-EntraExtensionManager.ps1 -Proxy "http://proxy.company.com:8080"

# With offline modules
.\AD-EntraExtensionManager.ps1 -OfflineModulesPath "C:\OfflineModules"
🔧 Configuration
Optional: Create a Configuration File
The tool auto-detects your domain, but you can specify default OUs by creating a config.json:
json{
    "DefaultOUs": [
        "OU=Computers,DC=yourdomain,DC=com",
        "OU=Workstations,DC=yourdomain,DC=com",
        "OU=Servers,DC=yourdomain,DC=com"
    ]
}
📖 Usage Guide
Tab 1: Set AD Attributes

Select an OU using the "Select OU" button
Choose which extension attribute to modify (1-15)
Enter the desired value
Load computers and review the list
Select computers to update
Click "Apply to Selected"

Tab 2: Sync to Entra ID

Add one or more OUs to sync
Select attributes to synchronize
Connect to Microsoft Graph (sign in required)
Enable/disable Preview Mode
Click "Start Sync"
Export results to CSV

🛠️ Advanced Features
Command-Line Parameters

-ConfigFile: Path to custom configuration file
-Proxy: HTTP/HTTPS proxy URL for restricted networks
-OfflineModulesPath: Path to offline PowerShell modules

For Air-Gapped Environments

Download required modules on an internet-connected machine:

powershellSave-Module -Name Microsoft.Graph.Authentication -Path C:\OfflineModules
Save-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Path C:\OfflineModules

Copy to target machine and run with -OfflineModulesPath

📊 Use Cases

Asset Management: Tag computers with location, department, or project codes
Software Deployment: Mark computers for specific software installations
Compliance Tracking: Flag computers for audit or compliance status
Cost Center Assignment: Associate computers with financial codes
Maintenance Windows: Define update schedules per computer group

🤝 Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

Fork the repository
Create your feature branch (git checkout -b feature/AmazingFeature)
Commit your changes (git commit -m 'Add some AmazingFeature')
Push to the branch (git push origin feature/AmazingFeature)
Open a Pull Request


Graph sync requires Azure AD Connect for device object presence
Extension attributes are limited to 1024 characters
Large OUs (>1000 computers) may take time to load

💡 Tips

Always use Preview Mode for first-time operations
Document your attribute usage standards
Run sync operations during off-peak hours for large environments
Export results for audit trails and compliance



Open an Issue
Check existing issues before creating new ones
Include error messages and PowerShell version in bug reports
