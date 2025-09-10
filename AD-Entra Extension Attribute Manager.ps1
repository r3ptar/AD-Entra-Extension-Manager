<#
.SYNOPSIS
  Unified tool for managing AD extension attributes and syncing them to Entra ID devices.

.DESCRIPTION
  - GUI interface for selecting OUs and managing extension attributes
  - Set extension attributes on AD computer objects
  - Sync extension attributes from AD to Entra ID via Microsoft Graph
  - Support for both immediate updates and batch operations
  - Preview mode for testing changes before applying

.PARAMETER OfflineModulesPath
  Folder containing required modules if PSGallery is blocked

.PARAMETER Proxy
  Optional HTTP/HTTPS proxy URL

.PARAMETER ConfigFile
  Path to JSON configuration file with default settings

.EXAMPLE
  .\AD-EntraExtensionManager.ps1

.EXAMPLE
  .\AD-EntraExtensionManager.ps1 -ConfigFile ".\config.json"

.EXAMPLE
  .\AD-EntraExtensionManager.ps1 -Proxy "http://proxy.company.com:8080"

.NOTES
  Version: 1.0
  Author: Your Organization
  Requires: RSAT ActiveDirectory
            Microsoft.Graph.Authentication
            Microsoft.Graph.Identity.DirectoryManagement
  Graph permission needed: Device.ReadWrite.All
  
  For first-time setup, create a config.json file:
  {
    "DefaultOUs": [
      "OU=Computers,DC=yourdomain,DC=com",
      "OU=Workstations,DC=yourdomain,DC=com"
    ]
  }
#>

[CmdletBinding()]
param(
    [string]$OfflineModulesPath,
    [string]$Proxy,
    [string]$ConfigFile = ".\config.json"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices
Add-Type -AssemblyName System.Drawing

#region Configuration

$script:Config = @{
    DefaultOUs = @()
}

# Load configuration if available
if (Test-Path $ConfigFile) {
    try {
        $jsonConfig = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        if ($jsonConfig.DefaultOUs) {
            $script:Config.DefaultOUs = $jsonConfig.DefaultOUs
        }
        Write-Verbose "Loaded configuration from $ConfigFile"
    } catch {
        Write-Warning "Failed to load configuration file: $_"
    }
}

# Auto-detect domain if no config
if ($script:Config.DefaultOUs.Count -eq 0) {
    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $domainDN = "DC=" + $domain.Name.Replace(".", ",DC=")
        
        # Common OU patterns - adjust based on your standard
        $script:Config.DefaultOUs = @(
            "OU=Computers,$domainDN",
            "OU=Workstations,$domainDN",
            "OU=Servers,$domainDN"
        )
        
        # Filter to only existing OUs
        $existingOUs = @()
        foreach ($ou in $script:Config.DefaultOUs) {
            try {
                $test = [ADSI]"LDAP://$ou"
                if ($test.Path) { $existingOUs += $ou }
            } catch {
                # OU doesn't exist, skip it
            }
        }
        
        if ($existingOUs.Count -gt 0) {
            $script:Config.DefaultOUs = $existingOUs
        } else {
            # If no standard OUs found, just use domain root
            $script:Config.DefaultOUs = @($domainDN)
        }
        
        Write-Verbose "Auto-detected domain OUs: $($script:Config.DefaultOUs -join ', ')"
    } catch {
        Write-Warning "Could not auto-detect domain. Please configure DefaultOUs manually."
        $script:Config.DefaultOUs = @()
    }
}

#endregion

#region Bootstrap & Utilities

# Prefer TLS 1.2
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Set-SessionProxy {
    param([string]$ProxyUrl)
    if ([string]::IsNullOrWhiteSpace($ProxyUrl)) { return }
    try {
        $wp = New-Object System.Net.WebProxy($ProxyUrl)
        [System.Net.WebRequest]::DefaultWebProxy = $wp
        [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        Write-Verbose ("Session proxy set to {0}" -f $ProxyUrl)
    } catch {
        Write-Warning ("Failed to set session proxy: {0}" -f $_.Exception.Message)
    }
}

function Ensure-PSGallery {
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
        }
    } catch {
        try {
            Register-PSRepository -Name 'PSGallery' `
                -SourceLocation 'https://www.powershellgallery.com/api/v2' `
                -InstallationPolicy Trusted -ErrorAction Stop
        } catch {
            Write-Verbose ("PSGallery unavailable: {0}" -f $_.Exception.Message)
        }
    }
}

function Import-FromOffline {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$OfflinePath
    )
    $candidate = Join-Path $OfflinePath $Name
    if (Test-Path $candidate) {
        Import-Module $candidate -ErrorAction Stop
        return $true
    }
    return $false
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$OfflinePath
    )
    $present = Get-Module -ListAvailable -Name $Name
    if ($present) { 
        Import-Module $Name -ErrorAction Stop
        return $true
    }

    if ($OfflinePath) {
        if (Import-FromOffline -Name $Name -OfflinePath $OfflinePath) { return $true }
    }

    try {
        Ensure-PSGallery
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module $Name -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Ensure-Dependencies {
    param([string]$OfflinePath)
    
    $script:modulesLoaded = @{
        ActiveDirectory = $false
        Graph = $false
    }
    
    $script:modulesLoaded.ActiveDirectory = Ensure-Module -Name ActiveDirectory -OfflinePath $OfflinePath
    
    if (Ensure-Module -Name Microsoft.Graph.Authentication -OfflinePath $OfflinePath) {
        if (Ensure-Module -Name Microsoft.Graph.Identity.DirectoryManagement -OfflinePath $OfflinePath) {
            $script:modulesLoaded.Graph = $true
        }
    }
    
    return $script:modulesLoaded
}

function Connect-GraphIfNeeded {
    if (-not $script:modulesLoaded.Graph) {
        throw "Microsoft Graph modules not loaded."
    }
    
    $ctx = $null
    try { $ctx = Get-MgContext } catch {}
    if (-not $ctx -or -not ($ctx.Scopes -contains 'Device.ReadWrite.All')) {
        Connect-MgGraph -Scopes 'Device.ReadWrite.All' -NoWelcome
    }
    try { Select-MgProfile -Name 'v1.0' } catch {}
}

#endregion

#region AD Functions

function Get-AllOrganizationalUnits {
    $rootOUs = $script:Config.DefaultOUs
    $allOUs = @()

    if ($rootOUs.Count -eq 0) {
        Write-Warning "No root OUs configured. Please select OUs manually or configure DefaultOUs."
        return @()
    }

    foreach ($root in $rootOUs) {
        try {
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$root")
            $searcher.Filter = "(objectClass=organizationalUnit)"
            $searcher.SearchScope = "Subtree"
            $searcher.PropertiesToLoad.Add("distinguishedName") | Out-Null

            $results = $searcher.FindAll()
            foreach ($result in $results) {
                $dn = $result.Properties["distinguishedName"][0]
                $allOUs += $dn
            }

            # Add the root OU itself
            $allOUs += $root
        } catch {
            Write-Verbose "Failed to enumerate OUs from $root : $_"
        }
    }

    # If no OUs found from config, try to get all OUs from domain
    if ($allOUs.Count -eq 0) {
        try {
            $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $domainDN = "DC=" + $domain.Name.Replace(".", ",DC=")
            
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainDN")
            $searcher.Filter = "(objectClass=organizationalUnit)"
            $searcher.SearchScope = "Subtree"
            $searcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
            
            $results = $searcher.FindAll()
            foreach ($result in $results) {
                $dn = $result.Properties["distinguishedName"][0]
                $allOUs += $dn
            }
        } catch {
            Write-Warning "Failed to enumerate OUs from domain root"
        }
    }

    return $allOUs | Sort-Object
}

function Get-ComputersInOU {
    param (
        [string]$OU,
        [bool]$Recursive = $true
    )

    $computers = @()
    
    try {
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$OU")
        $searcher.Filter = "(objectClass=computer)"
        $searcher.SearchScope = if ($Recursive) { "Subtree" } else { "OneLevel" }
        $searcher.PropertiesToLoad.Add("distinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("name") | Out-Null
        $searcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null
        
        # Add all extension attributes to properties to load
        1..15 | ForEach-Object {
            $searcher.PropertiesToLoad.Add("extensionAttribute$_") | Out-Null
        }
        
        $results = $searcher.FindAll()
        foreach ($result in $results) {
            $computer = @{
                DistinguishedName = $result.Properties["distinguishedName"][0]
                Name = $result.Properties["name"][0]
                sAMAccountName = $result.Properties["sAMAccountName"][0]
            }
            
            # Get extension attributes
            1..15 | ForEach-Object {
                $attrName = "extensionAttribute$_"
                if ($result.Properties.Contains($attrName) -and $result.Properties[$attrName].Count -gt 0) {
                    $computer[$attrName] = $result.Properties[$attrName][0]
                } else {
                    $computer[$attrName] = $null
                }
            }
            
            $computers += [PSCustomObject]$computer
        }
    } catch {
        Write-Warning "Failed to get computers from $OU : $_"
    }
    
    return $computers
}

function Update-ComputerExtensionAttribute {
    param (
        [string]$ComputerDN,
        [string]$Attribute,
        [string]$Value
    )

    try {
        $computer = [ADSI]"LDAP://$ComputerDN"
        if ([string]::IsNullOrEmpty($Value)) {
            $computer.PutEx(1, $Attribute, $null)  # Clear attribute
        } else {
            $computer.Put($Attribute, $Value)
        }
        $computer.SetInfo()
        return $true
    } catch {
        Write-Warning "Failed to update $ComputerDN : $_"
        return $false
    }
}

#endregion

#region Graph Functions

function Get-MgDeviceBySid {
    param([string]$Sid)
    if ([string]::IsNullOrWhiteSpace($Sid)) { return $null }
    try {
        $f = "onPremisesSecurityIdentifier eq '$Sid'"
        Get-MgDevice -Filter $f -ConsistencyLevel eventual -Count ct -All | Select-Object -First 1
    } catch {
        return $null
    }
}

function Get-MgDeviceByName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    try {
        $n = $Name.TrimEnd('$').Replace("'","''")
        $f = "displayName eq '$n'"
        Get-MgDevice -Filter $f -ConsistencyLevel eventual -Count ct -All | Select-Object -First 1
    } catch {
        return $null
    }
}

function Sync-ComputerToEntra {
    param(
        [PSCustomObject]$Computer,
        [int[]]$Attributes = (1..15),
        [switch]$Preview
    )
    
    $result = @{
        Computer = $Computer.sAMAccountName
        Status = "Unknown"
        Details = ""
        EntraDevice = $null
    }
    
    # Build extension body
    $ext = @{}
    $hasValues = $false
    foreach ($i in $Attributes) {
        if ($i -lt 1 -or $i -gt 15) { continue }
        $prop = "extensionAttribute$i"
        $value = $Computer.$prop
        if ($value) {
            $ext[$prop] = $value
            $hasValues = $true
        } else {
            $ext[$prop] = $null
        }
    }
    
    if (-not $hasValues) {
        $result.Status = "Skipped"
        $result.Details = "No attribute values to sync"
        return [PSCustomObject]$result
    }
    
    # Find Entra device
    $mgDevice = $null
    
    # Try by SID first if we have it
    if ($Computer.PSObject.Properties['SID']) {
        $mgDevice = Get-MgDeviceBySid -Sid $Computer.SID.Value
    }
    
    # Fall back to name
    if (-not $mgDevice) {
        $mgDevice = Get-MgDeviceByName -Name $Computer.sAMAccountName
    }
    
    if (-not $mgDevice) {
        $result.Status = "NoMatch"
        $result.Details = "No matching Entra device found"
        return [PSCustomObject]$result
    }
    
    $result.EntraDevice = $mgDevice.Id
    
    if ($Preview) {
        $result.Status = "Preview"
        $result.Details = "Would update: " + ($ext | ConvertTo-Json -Compress)
        return [PSCustomObject]$result
    }
    
    # Perform update
    try {
        $body = @{ extensionAttributes = $ext }
        Update-MgDevice -DeviceId $mgDevice.Id -BodyParameter $body -ErrorAction Stop | Out-Null
        $result.Status = "Success"
        $result.Details = "Updated successfully"
    } catch {
        $result.Status = "Error"
        $result.Details = $_.Exception.Message
    }
    
    return [PSCustomObject]$result
}

#endregion

#region GUI

function Show-MainWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "AD-Entra Extension Attribute Manager"
    $form.Size = New-Object System.Drawing.Size(900, 700)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false

    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(870, 650)
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $form.Controls.Add($tabControl)

    # Tab 1: Set Attributes
    $setTab = New-Object System.Windows.Forms.TabPage
    $setTab.Text = "Set AD Attributes"
    $setTab.UseVisualStyleBackColor = $true
    $tabControl.TabPages.Add($setTab)

    # Set Tab Controls
    $setOULabel = New-Object System.Windows.Forms.Label
    $setOULabel.Text = "Selected OU:"
    $setOULabel.Location = New-Object System.Drawing.Point(10, 20)
    $setOULabel.AutoSize = $true
    $setTab.Controls.Add($setOULabel)

    $setOUTextBox = New-Object System.Windows.Forms.TextBox
    $setOUTextBox.Location = New-Object System.Drawing.Point(100, 20)
    $setOUTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $setOUTextBox.ReadOnly = $true
    $setTab.Controls.Add($setOUTextBox)

    $selectOUButton = New-Object System.Windows.Forms.Button
    $selectOUButton.Text = "Select OU"
    $selectOUButton.Location = New-Object System.Drawing.Point(610, 18)
    $selectOUButton.Size = New-Object System.Drawing.Size(100, 23)
    $setTab.Controls.Add($selectOUButton)

    # Attribute selection
    $attrLabel = New-Object System.Windows.Forms.Label
    $attrLabel.Text = "Attribute:"
    $attrLabel.Location = New-Object System.Drawing.Point(10, 60)
    $attrLabel.AutoSize = $true
    $setTab.Controls.Add($attrLabel)

    $attrComboBox = New-Object System.Windows.Forms.ComboBox
    $attrComboBox.Location = New-Object System.Drawing.Point(100, 60)
    $attrComboBox.Size = New-Object System.Drawing.Size(200, 20)
    $attrComboBox.DropDownStyle = 'DropDownList'
    1..15 | ForEach-Object { $attrComboBox.Items.Add("extensionAttribute$_") }
    $attrComboBox.SelectedIndex = 0
    $setTab.Controls.Add($attrComboBox)

    # Value field
    $valueLabel = New-Object System.Windows.Forms.Label
    $valueLabel.Text = "Value:"
    $valueLabel.Location = New-Object System.Drawing.Point(10, 100)
    $valueLabel.AutoSize = $true
    $setTab.Controls.Add($valueLabel)

    $valueTextBox = New-Object System.Windows.Forms.TextBox
    $valueTextBox.Location = New-Object System.Drawing.Point(100, 100)
    $valueTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $setTab.Controls.Add($valueTextBox)

    # Recursive checkbox
    $recursiveCheckBox = New-Object System.Windows.Forms.CheckBox
    $recursiveCheckBox.Text = "Include Sub-OUs"
    $recursiveCheckBox.Location = New-Object System.Drawing.Point(100, 140)
    $recursiveCheckBox.AutoSize = $true
    $recursiveCheckBox.Checked = $true
    $setTab.Controls.Add($recursiveCheckBox)

    # Computer list
    $compListLabel = New-Object System.Windows.Forms.Label
    $compListLabel.Text = "Computers to update:"
    $compListLabel.Location = New-Object System.Drawing.Point(10, 180)
    $compListLabel.AutoSize = $true
    $setTab.Controls.Add($compListLabel)

    $compListView = New-Object System.Windows.Forms.ListView
    $compListView.Location = New-Object System.Drawing.Point(10, 200)
    $compListView.Size = New-Object System.Drawing.Size(840, 300)
    $compListView.View = 'Details'
    $compListView.FullRowSelect = $true
    $compListView.GridLines = $true
    $compListView.CheckBoxes = $true
    $compListView.Columns.Add("Computer Name", 200)
    $compListView.Columns.Add("Current Value", 300)
    $compListView.Columns.Add("Distinguished Name", 330)
    $setTab.Controls.Add($compListView)

    # Load computers button
    $loadComputersButton = New-Object System.Windows.Forms.Button
    $loadComputersButton.Text = "Load Computers"
    $loadComputersButton.Location = New-Object System.Drawing.Point(720, 58)
    $loadComputersButton.Size = New-Object System.Drawing.Size(130, 23)
    $setTab.Controls.Add($loadComputersButton)

    # Select all/none buttons
    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = "Select All"
    $selectAllButton.Location = New-Object System.Drawing.Point(10, 510)
    $selectAllButton.Size = New-Object System.Drawing.Size(100, 23)
    $setTab.Controls.Add($selectAllButton)

    $selectNoneButton = New-Object System.Windows.Forms.Button
    $selectNoneButton.Text = "Select None"
    $selectNoneButton.Location = New-Object System.Drawing.Point(120, 510)
    $selectNoneButton.Size = New-Object System.Drawing.Size(100, 23)
    $setTab.Controls.Add($selectNoneButton)

    # Apply button
    $applyButton = New-Object System.Windows.Forms.Button
    $applyButton.Text = "Apply to Selected"
    $applyButton.Location = New-Object System.Drawing.Point(720, 510)
    $applyButton.Size = New-Object System.Drawing.Size(130, 30)
    $applyButton.BackColor = [System.Drawing.Color]::LightGreen
    $setTab.Controls.Add($applyButton)

    # Progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 550)
    $progressBar.Size = New-Object System.Drawing.Size(840, 20)
    $progressBar.Style = 'Continuous'
    $setTab.Controls.Add($progressBar)

    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Ready"
    $statusLabel.Location = New-Object System.Drawing.Point(10, 575)
    $statusLabel.Size = New-Object System.Drawing.Size(840, 20)
    $setTab.Controls.Add($statusLabel)

    # Tab 2: Sync to Entra
    $syncTab = New-Object System.Windows.Forms.TabPage
    $syncTab.Text = "Sync to Entra ID"
    $syncTab.UseVisualStyleBackColor = $true
    $tabControl.TabPages.Add($syncTab)

    # Sync Tab Controls
    $syncOULabel = New-Object System.Windows.Forms.Label
    $syncOULabel.Text = "OU(s) to Sync:"
    $syncOULabel.Location = New-Object System.Drawing.Point(10, 20)
    $syncOULabel.AutoSize = $true
    $syncTab.Controls.Add($syncOULabel)

    $syncOUListBox = New-Object System.Windows.Forms.ListBox
    $syncOUListBox.Location = New-Object System.Drawing.Point(10, 40)
    $syncOUListBox.Size = New-Object System.Drawing.Size(600, 100)
    $syncOUListBox.SelectionMode = 'MultiExtended'
    $syncTab.Controls.Add($syncOUListBox)

    $addOUButton = New-Object System.Windows.Forms.Button
    $addOUButton.Text = "Add OU"
    $addOUButton.Location = New-Object System.Drawing.Point(620, 40)
    $addOUButton.Size = New-Object System.Drawing.Size(100, 23)
    $syncTab.Controls.Add($addOUButton)

    $removeOUButton = New-Object System.Windows.Forms.Button
    $removeOUButton.Text = "Remove Selected"
    $removeOUButton.Location = New-Object System.Drawing.Point(620, 70)
    $removeOUButton.Size = New-Object System.Drawing.Size(100, 23)
    $syncTab.Controls.Add($removeOUButton)

    $clearOUButton = New-Object System.Windows.Forms.Button
    $clearOUButton.Text = "Clear All"
    $clearOUButton.Location = New-Object System.Drawing.Point(620, 100)
    $clearOUButton.Size = New-Object System.Drawing.Size(100, 23)
    $syncTab.Controls.Add($clearOUButton)

    # Attributes to sync
    $syncAttrLabel = New-Object System.Windows.Forms.Label
    $syncAttrLabel.Text = "Attributes to Sync:"
    $syncAttrLabel.Location = New-Object System.Drawing.Point(10, 150)
    $syncAttrLabel.AutoSize = $true
    $syncTab.Controls.Add($syncAttrLabel)

    $syncAttrCheckedList = New-Object System.Windows.Forms.CheckedListBox
    $syncAttrCheckedList.Location = New-Object System.Drawing.Point(10, 170)
    $syncAttrCheckedList.Size = New-Object System.Drawing.Size(200, 250)
    $syncAttrCheckedList.CheckOnClick = $true
    1..15 | ForEach-Object { 
        $syncAttrCheckedList.Items.Add("extensionAttribute$_", $true)
    }
    $syncTab.Controls.Add($syncAttrCheckedList)

    # Graph connection status
    $graphStatusLabel = New-Object System.Windows.Forms.Label
    $graphStatusLabel.Text = "Graph Status: Not Connected"
    $graphStatusLabel.Location = New-Object System.Drawing.Point(250, 170)
    $graphStatusLabel.Size = New-Object System.Drawing.Size(300, 20)
    $graphStatusLabel.ForeColor = [System.Drawing.Color]::Red
    $syncTab.Controls.Add($graphStatusLabel)

    $connectGraphButton = New-Object System.Windows.Forms.Button
    $connectGraphButton.Text = "Connect to Graph"
    $connectGraphButton.Location = New-Object System.Drawing.Point(250, 200)
    $connectGraphButton.Size = New-Object System.Drawing.Size(150, 30)
    if (-not $script:modulesLoaded.Graph) {
        $connectGraphButton.Enabled = $false
        $connectGraphButton.Text = "Graph Modules Missing"
    }
    $syncTab.Controls.Add($connectGraphButton)

    # Preview checkbox
    $previewCheckBox = New-Object System.Windows.Forms.CheckBox
    $previewCheckBox.Text = "Preview Mode (Don't write changes)"
    $previewCheckBox.Location = New-Object System.Drawing.Point(250, 250)
    $previewCheckBox.Size = New-Object System.Drawing.Size(250, 20)
    $previewCheckBox.Checked = $true
    $syncTab.Controls.Add($previewCheckBox)

    # Sync button
    $syncButton = New-Object System.Windows.Forms.Button
    $syncButton.Text = "Start Sync"
    $syncButton.Location = New-Object System.Drawing.Point(250, 280)
    $syncButton.Size = New-Object System.Drawing.Size(150, 40)
    $syncButton.BackColor = [System.Drawing.Color]::LightBlue
    $syncButton.Enabled = $false
    $syncTab.Controls.Add($syncButton