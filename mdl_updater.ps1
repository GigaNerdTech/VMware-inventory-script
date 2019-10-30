# MDL Updater

if (Get-Module -ListAvailable -Name VMware.VimAutomation.Core) {
    Import-Module VMware.VimAutomation.Core
}
else {
    if (Get-PSRepository | Where-Object { $_ -match "example-nuget" }) {
        Install-Module -Name VMware.PowerCLI -Repository "example-nuget" -AllowClobber
        Import-Module VMware.VimAutomation.Core
    }
    else {
        Register-PSRepository -Name 'example-nuget' -SourceLocation 'https://nuget.example.int/nuget' -InstallationPolicy Trusted
        Install-Module -Name VMware.PowerCLI -Repository "example-nuget" -AllowClobber
        Import-Module VMware.VimAutomation.Core
    }
}
Import-Module ActiveDirectory

# Function to get folder paths to check for templates
function folder_path {
    Param($folder)
    $parent = $folder | Select -ExpandProperty Parent
    if ($parent -eq $null) {
        return $folder.Name
    }
    else {
        $new_path = folder_path $parent
        return  $new_path + "\" + $folder.Name
    }
}
# Start Transcript
$TranscriptFile = "\\networkshare\Powershell\PSLogs\MDLUpdater_$(get-date -f MMddyyyyHHmmss).txt"
$start_time = Get-Date
Start-Transcript -Path $TranscriptFile
$vhosts = @("vcenter-dal.mhd.com","vcenter-man.mhd.com","examplevc01","examplevc02")

$server_list_file = "\\networkshare\Combined Server List_$(get-date -f MMddyyyyHHmmss).xlsx"

Copy-Item "\\networkshare\Combined Server List.xlsx" -Destination $server_list_file -Force

$blank_rows = @()
$vm_inventory = @{}
$vm_list = @()
$mdl_inventory = @()
$physical_servers = @()
$hosts_with_punctuation = @()
$vms_not_in_mdl = @()
$vms_not_in_vcenter = @()
$lowercase_hosts = @()
$vp_inventory = @{}
$vp_list = @()
$template_inventory = @{}
$template_list = @()
$marked_templates = @()
$decomm_servers = @()

# Collect vSphere credentials
Write-Output "`n`nvSphere credentials:`n"
$vsphere_user = Read-Host -Prompt "Enter the user for the vCenter host"
$vsphere_pwd = Read-Host -Prompt "Enter the password for connecting to vSphere: " -AsSecureString
$vsphere_creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $vsphere_user,$vsphere_pwd -ErrorAction Stop


foreach ($vcenter_host in $vhosts) {
    # Connect to vCenter
    Connect-VIServer -Server $vcenter_host -Credential $vsphere_creds

    # Get All VMs
     Write-Host "Gathering all VMs in $vcenter_host..."
    $vm_collection = Get-VM -Server $vcenter_host

    # Get all VM physical hosts
    Write-Host "Gathering all physical hosts for vCenter $vcenter_host..."
    $vp_collection = Get-VMHost -Server $vcenter_host

    # Get all templates for comparison
    Write-Host "Gathering templates from vCenter $vcenter_host..."
    $template_collection = Get-Template -Server $vcenter_host

    $vm_inventory[$vcenter_host] = $vm_collection
    $vp_inventory[$vcenter_host] = $vp_collection
    $template_inventory[$vcenter_host] = $template_collection
}

Write-Host "Inventorying template names..."
foreach ($vcenter_host in $vhosts) {
    $template_inventory[$vcenter_host] | ForEach-Object {
        $template_list += $_.Name
    }
}

Write-Host "Getting current MDL inventory..."
# Open Excel

Write-Host "Opening Excel..."
# $erroractionpreference = "SilentlyContinue"
$excel_object = New-Object -comobject Excel.Application
$excel_object.visible = $True 

# Open existing Excel file
$excel_workbook = $excel_object.Workbooks.Open($server_list_file)

# Get Servers MDL Worksheet
$worksheet1 = $excel_workbook.Worksheets.Item('Servers MDL')

# Add new header column
$worksheet1.Cells.Item(1,22) = "MDL Script Comment"
$worksheet1.Cells.Item(1,22).Font.Bold = $true
$worksheet1.Cells.Item(1,22).Interior.ColorIndex = 7


# Get number of rows
$excel_range = $worksheet1.UsedRange
$row_count = $excel_range.Rows.Count

# Loop through all servers, get names
Write-Host "Processing MDL list..."

# These attempt to greatly speed up processing
$mdl_list = @($worksheet1.Columns(1).Value2)
$virtual_state = @($worksheet1.Columns(6).Value2)
$mdl_comments = @($worksheet1.Columns(22).Value2)

for ($i = 2; $i -le $row_count; $i++) {
    $current_host = $mdl_list[$i - 2]
    if ($current_host -match "\s|\(|\)|\.") {
        Write-Host "Server Name $current_host has whitespace or punctuation in it!"
        $hosts_with_punctuation += $current_host
    }
    if ([string]::IsNullOrWhiteSpace($current_host)) {
        Write-Host "Server name is blank on row $i!"
        $blank_rows += $i.ToString()
    }
    if ($current_host -cnotmatch $current_host.ToUpper()) {
        Write-Host "$current_host is not all uppercase!"
        $lowercase_hosts += $current_host
    }
#    if ($worksheet1.Cells.Item($i,6).Text -match "VM|virtual") {
     if ($virtual_state[$i-2] -match "VM|Virtual") {
        $mdl_inventory += $current_host
        Write-Host ("Found server name " + $current_host)
    }
    else {
        Write-Host ("Server " + $current_host + " marked as physical, skipping.")
        $physical_servers += $current_host
    }
    if ($mdl_comments[$i-2] -match "decomm|duplicate") {
        Write-Host "Server $current_host already marked for decomm, skipping!" -ForegroundColor Yellow
        $decomm_servers += $current_host
    }
        
}
$vm_properties =@{}
# Run compare of MDL with VMs
Write-Host "Looking for VMs in vCenter not in MDL..."
foreach ($vcenter in $vhosts) {
    $vm_inventory[$vcenter] | ForEach-Object {
        $current_vm = $_
        
        $current_vm_name = $current_vm.Name.ToString()

        $vm_properties[$current_vm_name] = $current_vm
        if (($mdl_inventory | ForEach {"$($_)"}) -contains $current_vm_name) {
            Write-Host "Current VM $current_vm_name found in MDL inventory." -ForegroundColor Green
        }
        else {
            Write-Host "Current VM $current_vm_name not found in MDL inventory." -ForegroundColor Red
            $vms_not_in_mdl += $current_vm_name
       }
       $vm_list += $current_vm_name
    }
}

# Run compare of VMs with MDL
Write-Host "Looking for VMs in MDL that are not in vCenter..."
foreach ($vm in $mdl_inventory) {
    if (($decomm_servers | ForEach {"$($_)"}) -contains $vm.ToString()) {
        Write-Host "Server $vm marked for decomm, skipping!" -ForegroundColor Yellow
    }
    elseif (($vm_list | ForEach {"$($_)"}) -contains $vm.ToString()) {
        Write-Host "VM $vm found in vCenter." -ForegroundColor Green
    }
    else {
        Write-Host "VM $vm not found in vCenter." -ForegroundColor Red
        $vms_not_in_vcenter += $vm
    }
}

# Check all current entries for template status
Write-Host "Checking for templates..."
for($i=2; $i -le $row_count; $i++) {
    $entry = $mdl_list[$i-2]
    Write-Host "Checking $entry for template status..."
    if (($template_list | ForEach {"$($_)" } ) -contains $entry.ToString()) {
        Write-Host "$entry is a template!" -ForegroundColor Green
        $marked_templates += $entry.ToString()
    }
<#    $path = (folder_path ($vm_properties[$entry.ToString()].Folder)).ToString()
    Write-Host "Found a folder path of $path"
    if ($path -match "template") {
        Write-Host "$entry is a template!" -ForegroundColor Green
        $marked_templates += $entry.ToString()
    }#>
    if ($entry -match "template|gold|image|master|base|img") {
        Write-Host "$entry is a template!" -ForegroundColor Green
        $marked_templates += $entry.ToString()
    }
}

$unresponsive_physical_servers = @()
$physicals_that_are_really_vms = @()

# Ping physical host to see if still responsive
Write-Host "Testing physical servers..."
foreach ($phost in $physical_servers) {
    Write-Host "Testing physical server $phost..."
    if (($decomm_servers | ForEach {"$($_)"}) -contains $phost) {
        Write-Host "Server $phost marked for decomm, skipping!" -ForegroundColor DarkYellow
    }
    elseif (($vm_list | Foreach { "$($_)" }) -contains $phost) {
        Write-Host "$phost found in VM inventory...changing to Virtual!" -ForegroundColor Yellow
        $physicals_that_are_really_vms += $phost
    }
    elseif (Test-Connection -ComputerName ($phost.ToString()) -Count 1 -Quiet) {
        Write-Host "Physical server $phost responded!" -ForegroundColor Green
    }
    else {
        Write-Host "Physical server $phost did not respond!" -ForegroundColor Red
        $unresponsive_physical_servers += $phost.ToString()
    }
        
}
# Update MDL with VMs not found in MDL
$new_rows = $row_count + 1
$vm_duplicate_rows = @{}
$template_count = 0
Write-Host "Adding hosts not in MDL..."
# Adding hash of arrays for Excel rows

$found_row = 0
$host_found = $false
$template_host = ""
foreach ($vm in ($vms_not_in_mdl | Sort)) {
    
    # Initialize array

    Write-Host "Adding host $vm..."
    $vm_name = $vm.ToString()
    if (($template_list | ForEach { "$($_)" }) -match $vm_name) {
        Write-Host "VM $vm is a template! Marking..."
        $worksheet1.Cells.Item($new_rows,6) = "Template"
        $template_count++
    }
    else {
        $worksheet1.Cells.Item($new_rows,6) = "Virtual"

    }
    $worksheet1.Cells.Item($new_rows,1) = $vm.ToString().ToUpper()

    $worksheet1.Cells.Item($new_rows,7) = (Get-Datacenter -VM $vm_properties[$vm.ToString()]).Name
    $worksheet1.Cells.Item($new_rows,8) = ((Get-Cluster -VM $vm_properties[$vm.ToString()]).Name + " Cluster")
    $worksheet1.Cells.Item($new_rows,9) = "example"
    $worksheet1.Cells.Item($new_rows,10) = (($vm_properties[$vm.ToString()] | Get-View -Property @("Name","Config.GuestFullName","Guest.GuestFullName") | Select -Property @{N="Running OS";E={$_.Config.GuestFullName}} | Out-String) -replace "Running OS","" -replace "----------","" -replace "or later","").Trim()
    if ($worksheet1.Cells.Item($new_rows,10).Text -match "Windows") {
        $win_license = ( Get-WmiObject SoftwareLicensingProduct -ComputerName $vm_name | Select -ExpandProperty LicenseFamily | Select -First 1)
        $worksheet1.Cells.Item($new_rows,10) = ($worksheet1.Cells.Item($new_rows,10).Text + " " + $win_license)
    }
    $worksheet1.Cells.Item($new_rows,21) = ($vm_properties[$vm.ToString()].Notes).ToString()
    $worksheet1.Cells.Item($new_rows,22) =  "Added by MDL Script!"
    

        # Check for SQL
    if (($worksheet1.Cells.Item($new_rows,10).Text -match "Windows Server") -and (($worksheet1.Cells.Item($new_rows,11).Text -match "Y|YES|SQL") -or ([string]::IsNullOrWhiteSpace($worksheet1.Cells.Item($new_rows,11).Text)) -and $worksheet1.Cells.Item($new_rows,22).Text -notmatch "decomm") -and ($vm_name -notmatch "EPC-...-HS")) { 
            # Determine WMI namespace for current computer
            $license = ""
            $version = ""
            Write-Host "Checking $vm_name for SQL..."
         $comp_mgmt_nsp = (Get-WmiObject -ComputerName $vm_name -Namespace "root\microsoft\sqlserver" -Class __NAMESPACE -ErrorAction SilentlyContinue |
                Where-Object {$_.Name -like "ComputerManagement*"} |
                Select-Object Name |
                Sort-Object Name -Descending |
                Select-Object -First 1).Name


        $comp_mgmt_nsp = "root\microsoft\sqlserver\" + $comp_mgmt_nsp

        # Get SQL Server license type using WMI
        $license_object = Get-WmiObject -ComputerName $vm_name -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "SKUNAME"
            } |
            Select-Object @{Name = "ComputerName"; Expression = { $current_host }},
            @{Name = "PropertyValue"; Expression = {
                 $_.PropertyStrValue}  
            } 
            if (($license_object | Get-Unique).Count -gt 1) {
                $license_object = ($license_object | Where-Object { $_.PropertyValue -match "Standard" -or $_.PropertyValue -match "Enterprise" }) 
            }
    
        # Get SQL version using WMI        
        $version_object = Get-WmiObject -ComputerName $vm_name -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "VERSION"
            } -ErrorAction SilentlyContinue |
            Select-Object @{Name = "ComputerName"; Expression = { $vm_name }},
            @{Name = "SQLVersion"; Expression = {
                 $_.PropertyStrValue}      
            } | Select-Object -First 1

    
        # Check to see if version exists, if so, set it properly
        if (Get-Member -InputObject $version_object -name "SQLVersion" -MemberType Properties) {
            $version = ($version_object | Select -ExpandProperty SQLVersion).ToString()
        }
        else {
                $version = "None"
        }

        # Set license type
        $license = ( $license_object | Select-Object -First 1 | Select -ExpandProperty PropertyValue).ToString()
        if ($license -match "Express|Standard|Enterprise") {
            Write-Host "MS SQL Server found on $vm_name! License: $license Version: $version" -ForegroundColor Green
            $worksheet1.Cells.Item($new_rows,11) = "Y"
            $worksheet1.Cells.Item($new_rows,21) = ($worksheet1.Cells.Item($new_rows,21).Text + " " + "SQL License: $license SQL Version: $version") 

        }
        else {
            Write-Host "SQL Not found on $vm_name!" -ForegroundColor Red
            $worksheet1.Cells.Item($new_rows,11) = "N"

        }
     }
  
     elseif ($worksheet1.Cells.Item($new_rows,10).Text -notmatch "Windows Server") {

        $worksheet1.Cells.Item($new_rows,11) = "N"
     }
     else { 
        # Do nothing
   
    }
    # Find template for new host
    if ((($vm_name | Select-String -Pattern "(.+?)\d{1,3}").Matches.Groups[1].Value -ne $template_host) -and (-not ([string]::IsNullOrEmpty(($vm_name | Select-String -Pattern "(.+?)\d{1,3}").Matches.Groups[1].Value)))) {
            Write-Host "Need new template!"
            $found_row = 0
            $host_found = $false
            $template_host = ""
            $template_host = ($vm_name | Select-String -Pattern "(.+?)\d{1,3}$").Matches.Groups[1].Value

            if ((-not [string]::IsNullOrWhitespace($template_host))) {
                $host_found = $true
                $found_row = [array]::IndexOf($mdl_list,($mdl_list | Select-String -Pattern "$template_host\d+").Matches.Groups[0].Value) + 2
                if ([string]::IsNullOrEmpty($worksheet1.Cells.Item($found_row,3).Text) -or ($worksheet1.Cells.Item($found_row,3).Text -match [Regex]::Escape('**?**'))) {
                    $found_row++
                }
                Write-Host "Found a template host $template_host for $vm_name on row $found_row!" -ForegroundColor Green
                if (($found_row -le 1) -or ($worksheet1.Cells.Item($found_row, 3).Text -match [Regex]::Escape('**?**'))) {
                    $found_row = 0
                    $host_found = $false
                    $template_host = ""
                    Write-Host "Not applying template!" -ForegroundColor Red
                }
            }
        }

        
     # Go through each cell

     if (-not ([string]::IsNullOrEmpty($template_host)) -and ($vm_name -match "\d+")) {
        for ($j=2; $j -le 21; $j++) {

            if (($j -ne 6) -and ($j -ne 7) -and ($j -ne 8) -and ($j -ne 10) -and ($j -ne 11)) {
            # If the duplicate original is blank or marked as **?**, check to see if template found if it's not column 1, 6, 7, 8, or 10
            
                if (($found_row -ge 2) -and ($worksheet1.Cells.Item($found_row,$j).Text -notmatch [Regex]::Escape('**?**'))) {

                    #  if there is a "template" host in the MDL already, set the new row to that cell's value
                    if (-not ([string]::IsNullOrWhiteSpace($worksheet1.Cells.Item($found_row,$j).Text)) -and ($worksheet1.Cells.Item($new_rows,$j).Text -notmatch [Regex]::Escape('**?**'))) {
                        
                        $worksheet1.Cells.Item($new_rows,$j) = $worksheet1.Cells.Item($found_row,$j).Text
                    }              
                }
         }
      }
    }
    $worksheet1.Range($worksheet1.Cells($new_rows,1),$worksheet1.Cells($new_rows,22)).Interior.ColorIndex = 4
    $vm_duplicate_rows[$vm_name] = $new_rows
    $new_rows++
}

$vp_count = 0


# Add vCenter physical hosts

Write-Host "Adding vcenter physical hosts..."
foreach ($vcenter in $vhosts) {
    $vp_inventory[$vcenter] | ForEach-Object {

        $vp = $_
        $vp_name = $vp.Name
        if (($physical_servers | ForEach { "$($_)" }) -notcontains ($vp.Name -replace "\.example\.int|\.mhd\.com","").ToUpper()) {
        Write-Host "Adding vCenter physical server $vp_name..."
        $worksheet1.Cells.Item($new_rows,1) = ($vp.Name).ToUpper()
        $worksheet1.Cells.Item($new_rows,2) = "1"
        $worksheet1.Cells.Item($new_rows,3) = "VMware"
        $worksheet1.Cells.Item($new_rows,4) = "VMware ESXi"
        $worksheet1.Cells.Item($new_rows,6) = "Physical"
        $worksheet1.Cells.Item($new_rows,7) = (Get-Datacenter -VMHost $vp).Name
        $worksheet1.Cells.Item($new_rows,8) = ((Get-Cluster -VMHost $vp).Name + " Cluster")
        $worksheet1.Cells.Item($new_rows,9) = ""
        $worksheet1.Cells.Item($new_rows,10) = "ESXi"
        $worksheet1.Cells.Item($new_rows,12) = "Server Team"
        $worksheet1.Cells.Item($new_rows,13) = "Mananger"
        $worksheet1.Cells.Item($new_rows,22) = "Added by MDL Script!"
        $worksheet1.Range($worksheet1.Cells($new_rows,1),$worksheet1.Cells($new_rows,22)).Interior.ColorIndex = 4
        $new_rows++
        $vp_count++
    }
    }
}

# Process spreadsheet
$found_row = 0
$host_found = $false
$template_host = ""

for ($i=2; $i -le $row_count; $i++) {
     $current_host = $worksheet1.Cells.Item($i,1).Text
    Write-Host ("Working on host " + $current_host)

    if (($decomm_servers | ForEach {"$($_)"}) -contains $vm.ToString()) {
        Write-Host "Server $current_host marked for decomm, skipping!" -ForegroundColor DarkYellow
        continue
    }
    if ($worksheet1.Cells.Item($i,22).Text -match "Modified on") {
        $modified_date = [datetime]::ParseExact((Select-String -InputObject $worksheet1.Cells.Item($i,22).Text -Pattern "Modified on (\d\d/\d\d/\d\d\d\d)").Matches.Groups[1].Value,"MM/dd/yyyy",$null)
        $difference = New-TimeSpan -end $start_time -start $modified_date

        if ($difference.TotalDays -lt 30) {
            Write-Host "$current_host last modified on $modified_date . Skipping!" -ForegroundColor DarkGreen
            continue
        }
    }
   # Mark MDL spreadsheet with VMs not in vCenter
    if (($vms_not_in_vcenter | Foreach { "$($_)"}) -contains $current_host) {
        $worksheet1.Cells.Item($i,22) = "Not present in vCenter, consider decomm!"
        Write-Host ($current_host + " Not present in vCenter, consider decomm!")
        $worksheet1.Range($worksheet1.Cells($i,1),$worksheet1.Cells($i,22)).Interior.ColorIndex = 3
        
    }
    # Mark MDL spreadsheet with unresponsive physical servers
    if (($unresponsive_physical_servers | Foreach { "$($_)"}) -contains $current_host) {
        $worksheet1.Cells.Item($i,22) = "Physical server not responding, consider decomm!"
        Write-Host ($current_host + " Physical server not responding, consider decomm!")
        $worksheet1.Range($worksheet1.Cells($i,1),$worksheet1.Cells($i,22)).Interior.ColorIndex = 6
    }



       
    # Check for template
    if (($marked_templates | ForEach { "$($_)" }) -contains $current_host) {
        Write-Host "$current_host is a template. Marking!"
        $worksheet1.Cells.Item($i,6) = "Virtual Template"
    }

    # Check for SQL
    if ((($physicals_that_are_really_vms | ForEach { "$($_)" }) -notcontains $current_host) -and ($worksheet1.Cells.Item($i,10).Text -match "Windows Server") -and (($worksheet1.Cells.Item($i,11).Text -match "Y|YES|SQL") -or ([string]::IsNullOrWhiteSpace($worksheet1.Cells.Item($i,11).Text)) -and $worksheet1.Cells.Item($i,22).Text -notmatch "decomm") -and ($current_host -notmatch "EPC-...HS")) { 
            # Determine WMI namespace for current computer
            $license = ""
            $version = ""
            Write-Host "Checking $current_host for SQL..."
         $comp_mgmt_nsp = (Get-WmiObject -ComputerName $current_host -Namespace "root\microsoft\sqlserver" -Class __NAMESPACE -ErrorAction SilentlyContinue |
                Where-Object {$_.Name -like "ComputerManagement*"} |
                Select-Object Name |
                Sort-Object Name -Descending |
                Select-Object -First 1).Name


        $comp_mgmt_nsp = "root\microsoft\sqlserver\" + $comp_mgmt_nsp

        # Get SQL Server license type using WMI
        $license_object = Get-WmiObject -ComputerName $current_host -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "SKUNAME"
            } |
            Select-Object @{Name = "ComputerName"; Expression = { $current_host }},
            @{Name = "PropertyValue"; Expression = {
                 $_.PropertyStrValue}  
            } 
            if (($license_object | Get-Unique).Count -gt 1) {
                $license_object = ($license_object | Where-Object { $_.PropertyValue -match "Standard" -or $_.PropertyValue -match "Enterprise" }) 
            }
    
        # Get SQL version using WMI        
        $version_object = Get-WmiObject -ComputerName $current_host -Namespace $comp_mgmt_nsp -Class "SqlServiceAdvancedProperty" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.ServiceName -like "MSSQL*" -and
                $_.PropertyName -eq "VERSION"
            } -ErrorAction SilentlyContinue |
            Select-Object @{Name = "ComputerName"; Expression = { $current_host }},
            @{Name = "SQLVersion"; Expression = {
                 $_.PropertyStrValue}      
            } | Select-Object -First 1

    
        # Check to see if version exists, if so, set it properly
        if (Get-Member -InputObject $version_object -name "SQLVersion" -MemberType Properties) {
            $version = ($version_object | Select -ExpandProperty SQLVersion).ToString()
        }
        else {
                $version = "None"
        }

        # Set license type
        $license = ( $license_object | Select-Object -First 1 | Select -ExpandProperty PropertyValue).ToString()
        if ($license -match "Express|Standard|Enterprise") {
            Write-Host "MS SQL Server found on $current_host! License: $license Version: $version" -ForegroundColor Green
            $worksheet1.Cells.Item($i,11) = "Y"
            $worksheet1.Cells.Item($i,21) = ($worksheet1.Cells.Item($i,21).Text + " " + "SQL License: $license SQL Version: $version") 
            if ($vm_duplicate_rows[$current_host] -ne $null) {
                $worksheet1.Cells.Item($vm_duplicate_rows[$current_host],11) = "Y"
                $worksheet1.Cells.Item($vm_duplicate_rows[$current_host],21) = ($worksheet1.Cells.Item($vm_duplicate_rows[$current_host],21).Text + " " + "SQL License: $license SQL Version: $version") 
            }  
        }
        else {
            Write-Host "SQL Not found on $current_host!" -ForegroundColor Red
            $worksheet1.Cells.Item($i,11) = "N" 
            if ($vm_duplicate_rows[$current_host] -ne $null) {
                $worksheet1.Cells.Item($vm_duplicate_rows[$current_host],11) = "N"
            }  
        }
     }
  
     elseif ($worksheet1.Cells.Item($i,10).Text -notmatch "Windows Server") {
           if ($vm_duplicate_rows[$current_host] -ne $null) {
                $worksheet1.Cells.Item($vm_duplicate_rows[$current_host],11) = "N"
            }  
        $worksheet1.Cells.Item($i,11) = "N"
     }
     else { 
        # Do nothing
   
    }
    ######## Consistency checks ##########
    # VM->Virtual
    if ($worksheet1.Cells.Item($i,6).Text -match "VM") {
        $worksheet1.Cells.Item($i,6) = "Virtual"
    }
     # Remove M and D from cluster names
    if ($worksheet1.Cells.Item($i,8).Text -match "\sD|\sM") {
        $worksheet1.Cells.Item($i,8) = $worksheet1.Cells.Item($i,8).Text -replace "\sD|\sM",""
    }
    # Check for punctuation or whitespace
    if (($hosts_with_punctuation | Foreach { "$($_)"}) -contains $current_host) {
        Write-Host "$current_host contains punctutation or whitespace!"
        $current_comment = $worksheet1.Cells.Item($i,22).Text
        $worksheet1.Cells.Item($i,22) = ($current_comment + " Name contains punctuation or whitespace!")
    }
    # Check for lowercase names
    if (($lowercase_hosts | Foreach { "$($_)" }) -contains $current_host) {
        Write-Host "$current_host is not all uppercase! Changed!"
        $worksheet1.Cells.Item($i,1) = $current_host.ToUpper()
    }
    # Check for blank rows
    if ([string]::IsNullOrWhiteSpace($current_host)) {
        Write-Host "Encountered blank row on row $i! Marking in column 22!"
        $worksheet1.Cells.Item($i,22) = "Blank server name!"
    }
    # Update all VMs with correct OS, Datacenter, and notes
    if ($worksheet1.Cells.Item($i,6).Text -match "Virtual") {
            Write-Host "Updating VM $current_host with vCenter data..."
            $worksheet1.Cells.Item($i,7) = (Get-Datacenter -VM $current_host).Name
            $worksheet1.Cells.Item($i,8) = ((Get-Cluster -VM $current_host).Name + " Cluster")
            
            $worksheet1.Cells.Item($i,10) = (($vm_properties[$vm.ToString()] | Get-View -Property @("Name","Config.GuestFullName","Guest.GuestFullName") | Select -Property @{N="Running OS";E={$_.Config.GuestFullName}} | Out-String) -replace "Running OS","" -replace "----------","" -replace "or later","").Trim()
            if ($worksheet1.Cells.Item($i,10).Text -match "Windows") {
                $win_license = ( Get-WmiObject SoftwareLicensingProduct -ComputerName $current_host | Select -ExpandProperty LicenseFamily | Select -First 1)
                $worksheet1.Cells.Item($i,10) = ($worksheet1.Cells.Item($i,10).Text + " " + $win_license)
            }
            $worksheet1.Cells.Item($i,21) = (($vm_properties[$vm.ToString()].Notes).ToString() + $worksheet1.Cells.Item($i,21).Text)
    }
       # Correct virtual servers marked as physical
    if (($physicals_that_are_really_vms | ForEach { "$($_)" }) -contains $current_host) {
        Write-Host "$current_host is actually a VM, not physical! Marking current entry to replace with entry from vCenter!"
           # Find template for new host
    if ((($current_host | Select-String -Pattern "(.+?)\d{1,3}").Matches.Groups[1].Value -ne $template_host) -and (-not ([string]::IsNullOrEmpty(($current_host | Select-String -Pattern "(.+?)\d{1,3}").Matches.Groups[1].Value)))) {
            Write-Host "Need new template!"
            $found_row = 0
            $host_found = $false
            $template_host = ""
            $template_host = ($current_host | Select-String -Pattern "(.+?)\d{1,3}$").Matches.Groups[1].Value

            if ((-not [string]::IsNullOrWhitespace($template_host))) {
                $host_found = $true
                $found_row = [array]::IndexOf($mdl_list,($mdl_list | Select-String -Pattern "$template_host\d+").Matches.Groups[0].Value) + 2
                if ([string]::IsNullOrEmpty($worksheet1.Cells.Item($found_row,3).Text) -or ($worksheet1.Cells.Item($found_row,3).Text -match [Regex]::Escape('**?**'))) {
                    $found_row++
                }
                Write-Host "Found a template host $template_host for $current_host on row $found_row!" -ForegroundColor Green
                if (($found_row -le 1) -or ($worksheet1.Cells.Item($found_row, 3).Text -match [Regex]::Escape('**?**'))) {
                    $found_row = 0
                    $host_found = $false
                    $template_host = ""
                    Write-Host "Not applying template!" -ForegroundColor Red
                }
            }
        }

        
     # Go through each cell

     if (-not ([string]::IsNullOrEmpty($template_host)) -and ($current_host -match "\d+")) {
        for ($j=2; $j -le 21; $j++) {

            if (($j -ne 6) -and ($j -ne 7) -and ($j -ne 8) -and ($j -ne 10) -and ($j -ne 11)) {
            # If the duplicate original is blank or marked as **?**, check to see if template found if it's not column 1, 6, 7, 8, or 10
            
                if (($found_row -ge 2) -and ($worksheet1.Cells.Item($found_row,$j).Text -notmatch [Regex]::Escape('**?**'))) {

                    #  if there is a "template" host in the MDL already, set the new row to that cell's value
                    if (-not ([string]::IsNullOrWhiteSpace($worksheet1.Cells.Item($found_row,$j).Text)) -and ($worksheet1.Cells.Item($new_rows,$j).Text -notmatch [Regex]::Escape('**?**'))) {
                        
                        $worksheet1.Cells.Item($new_rows,$j) = $worksheet1.Cells.Item($found_row,$j).Text
                    }              
                }
         }
      }
    }

  
        $worksheet1.Range($worksheet1.Cells($i,1),$worksheet1.Cells($i,22)).Interior.ColorIndex = 8
        $current_comment = $worksheet1.Cells.Item($i,22).Text
        $worksheet1.Cells.Item($i,22) = ($current_comment + " Duplicate of entry marked physical that is really virtual!")
    }
    # timestamp the modification
    $datestamp = (Get-date).ToString("MM/dd/yyy")
    $current_comment = $worksheet1.Cells.Item($i,22).Text
    $worksheet1.Cells.Item($i,22) = ($current_comment + " Modified on $datestamp")

     
}
Write-Host "Disconnecting from vCenter..."
foreach ($vcenter_host in $vhosts) {
    Disconnect-VIServer -Server $vcenter_host -Confirm:$false
}
Write-Host "Formatting Excel..."
#Sort by server name
$sort_range = $worksheet1.Columns(1)
$sort_range.Sort($sort_range,1,$null,$null,1,$null,1,1) | Out-null


# Word wrap function, team, OS and app columns
$worksheet1.UsedRange.WrapText = $true

# Set AutoFilter for header
$worksheet1.UsedRange.AutoFilter($null, $null, $null, $null, $true) | Out-null

# Make borders consistent
$border_range = $worksheet1.UsedRange
$border_range.Borders.Item(12).Weight = 2
$border_range.Borders.Item(12).LineStyle = 1
$border_range.Borders.Item(12).ColorIndex = 1

$border_range.Borders.Item(11).Weight = 2
$border_range.Borders.Item(11).LineStyle = 1
$border_range.Borders.Item(11).ColorIndex = 1

$border_range.BorderAround(1,4,1)

# Save Excel
$excel_workbook.Save() | out-null

Write-Host "Cleaning up..."
# Quit Excel
$excel_workbook.Close | out-null
$excel_object.Quit() | out-null
# Remove variables
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel_object)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable excel_object | Out-null

Stop-Transcript

$end_time = Get-Date

$runtime = New-TimeSpan -end $end_time -start $start_time
# Send email summary report
Write-Host "Sending email report..."

# Generate email report
$username = $env:USERNAME
$email_user = (Get-ADUser -Properties EmailAddress -Identity $username | Select -ExpandProperty EmailAddress)

$email_list=@($email_user)
$subject = "MDL Excel Report"

$body = @("<h1>MDL Analysis Report</h1><br><br>")

# $body += "Number of servers in current inventory: $row_count`n`Number of VMs added: " + $vms_not_in_mdl.Count + "`nNumber of VMs marked as possible decomm: " +  $vms_not_in_vcenter.Count + "`nNumber of physical servers marked as virtual: " + $physicals_that_are_really_vms.Count + "`nNumber of physical servers that didn't respond: " + $unresponsive_physical_servers.Count + "`nLink to transcript: $TranscriptFile`nLink to new Excel: `"$server_list_file`"`nRun by: " + $env:USERNAME + "`nRuntime: " + $runtime.Minutes + " minutes`n`n`nKEY:`n`nGreen: Added by MDL script`nRed: VM not present in vCenter`nYellow: Unresponsive physical server`nCyan: Duplicate entry for virtual server")
$body += "`nLink to transcript: $TranscriptFile<br>Link to new Excel: `"$server_list_file`"<br>Run by: " + $env:USERNAME + "<br>Runtime: " + $runtime.TotalMinutes + " minutes<br><br><br><h2>Statistics</h2>"
$body += "<table border=`"3`"><thead><tr><th>Parameter</th><th>Value</th></tr></thead><tbody>"
$body += "<tr><td>Number of servers in current inventory</td><td>$row_count</td></tr>"
$body += ("<tr><td>Number of VMs added</td><td>" + $vms_not_in_mdl.Count + "</td></tr>")
$body += ("<tr><td>Number of ESXi physical hosts added</td><td>" + $vp_count + "</td></tr>")
$body += ("<tr><td>Number of VMs marked as possible decomm</td><td>" + $vms_not_in_vcenter.Count + "</td></tr>")
$body += ("<tr><td>Number of physical servers marked as virtual</td><td>" + $physicals_that_are_really_vms.Count + "</td></tr>")
$body += ("<tr><td>Number of physical servers that didn't respond</td><td>" + $unresponsive_physical_servers.Count + "</td></tr>")
$body += ("<tr><td>Number of entries marked as templates</td><td>" + $marked_templates.Count + "</td></tr>")
$body += ("<tr><td>Number of entries added as templates</td><td>" + $template_count + "</td></tr>")
$body += ("<tr><td>Number of hosts with punctuation or whitespace in the name</td><td>" + $hosts_with_punctuation.Count + "</td></tr>")
$body += ("<tr><td>Number of hosts with inconsistent case</td><td>" + $lowercase_hosts.Count + "</td></tr>")
$body += ("<tr><td>Number of blank rows</td><td>" + $blank_rows.Count + "</td></tr>")
$body += "</tbody></table>"

$body += "<br><br><h2>Color Key</h2><br><br>"
$body +=  "<table border=`"3`"><thead><tr><th>Color</th><th>Meaning</th></tr></thead><tbody>"
$body += "<tr bgcolor=`"red`"><td>RED</td><td>Virtual Machine marked for Decommission</td></tr>"
$body += "<tr bgcolor=`"yellow`"><td>YELLOW</td><td>Physical Machine marked for Decommission</td></tr>"
$body += "<tr bgcolor=`"cyan`"><td>CYAN</td><td>Unknown or physical machines found in virtual server inventory</td></tr>"
$body += "<tr bgcolor=`"green`"><td>GREEN</td><td>Entries added by MDL script</td></tr>"
$body += "<tr><td>WHITE</td><td>Original MDL entry</td></tr>"
$body += "</tbody></table>"

$MailMessage = @{
    To = $email_list
    From = "MDLReport<Donotreply@example.com>"
    Subject = $subject
    Body = ($body -join "<br/>")
    SmtpServer = "smtp.example.com"
    ErrorAction = "Stop"
}
Send-MailMessage @MailMessage -BodyAsHtml | out-null
