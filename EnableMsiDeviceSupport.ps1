Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Find-MSIDevices {
    # Base registry path for PCI devices
    $baseKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\PCI"
    
    # Initialize collections
    $msiDevices = @()
    $totalDevicesScanned = 0

    Write-Host "===== Scanning PCI Devices for MSI Configuration =====" -ForegroundColor Cyan
    Write-Host "Analyzing devices for MSI support..." -ForegroundColor Green

    # Get all PCI devices
    $devices = Get-ChildItem -Path $baseKey -Recurse -ErrorAction SilentlyContinue

    foreach ($device in $devices) {
        $totalDevicesScanned++

        try {
            # Get friendly device name and extract part after semicolon
            $deviceName = (Get-ItemProperty -Path $device.PSPath -Name "DeviceDesc" -ErrorAction SilentlyContinue).DeviceDesc
            if ($deviceName) {
                # Extract part after semicolon if it exists
                if ($deviceName -match ';') {
                    $deviceName = ($deviceName -split ';', 2)[1].Trim()
                }
                # Remove any @oemXX patterns as a fallback
                $deviceName = $deviceName -replace '@oem\d+(\.inf)?$', ''
                $deviceName = $deviceName.Trim()
            }
            if (-not $deviceName) { $deviceName = "Unknown Device" }

            # Check MSI support status
            $msiSupported = Get-ItemProperty -Path $device.PSPath -Name "MSISupported" -ErrorAction SilentlyContinue
            $msiStatus = if ($null -eq $msiSupported) { "Not Configured" } 
                        elseif ($msiSupported.MSISupported -eq 1) { "Enabled" } 
                        else { "Disabled" }

            # Prepare device information
            $deviceInfo = [PSCustomObject]@{
                DevicePath = $device.PSPath
                DeviceName = $deviceName
                MSIStatus = $msiStatus
                Action = $msiStatus  # Default action is current status
            }
            $msiDevices += $deviceInfo

            Write-Host "`n[Device]" -ForegroundColor Yellow
            Write-Host "Device: $deviceName" -ForegroundColor White
            Write-Host "Path: $($device.PSPath)" -ForegroundColor White
            Write-Host "MSI Status: $msiStatus" -ForegroundColor White
        }
        catch {
            Write-Host "Error processing device: $($device.PSPath)" -ForegroundColor Red
            Write-Host "Error Details: $_" -ForegroundColor DarkRed
        }
    }

    Write-Host "`nFound $totalDevicesScanned devices, $($msiDevices.Count) with MSI configuration options." -ForegroundColor Cyan
    return $msiDevices, $totalDevicesScanned
}

function Show-DeviceSelectionGUI {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Devices
    )

    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Configure MSI Support for Devices"
    $form.Size = New-Object System.Drawing.Size(800, 400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Create checkbox for showing unknown devices
    $checkBox = New-Object System.Windows.Forms.CheckBox
    $checkBox.Location = New-Object System.Drawing.Point(10, 10)
    $checkBox.Size = New-Object System.Drawing.Size(150, 20)
    $checkBox.Text = "Show Unknown Devices"
    $checkBox.Checked = $false

    # Create button to disable all unknown devices
    $disableUnknownButton = New-Object System.Windows.Forms.Button
    $disableUnknownButton.Location = New-Object System.Drawing.Point(170, 10)
    $disableUnknownButton.Size = New-Object System.Drawing.Size(150, 20)
    $disableUnknownButton.Text = "Disable for Unknown Devices"

    # Create a DataGridView for devices
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Location = New-Object System.Drawing.Point(10, 40)
    $dataGridView.Size = New-Object System.Drawing.Size(760, 270)
    $dataGridView.ColumnCount = 3
    $dataGridView.Columns[0].Name = "Device Name"
    $dataGridView.Columns[0].Width = 300
    $dataGridView.Columns[1].Name = "Current MSI Status"
    $dataGridView.Columns[1].Width = 150
    $dataGridView.Columns[2].Name = "Action"
    $dataGridView.Columns[2].Width = 150
    $dataGridView.AllowUserToAddRows = $false  # Prevent empty row

    # Add combo box column for actions
    $comboBoxColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $comboBoxColumn.Name = "Action"
    $comboBoxColumn.Items.AddRange(@("Not Configured", "Enabled", "Disabled"))
    $dataGridView.Columns.RemoveAt(2)
    $dataGridView.Columns.Add($comboBoxColumn)

    # Initialize filtered indices
    $script:filteredIndices = @()

    # Function to populate DataGridView based on checkbox state
    function Update-DataGridView {
        $dataGridView.Rows.Clear()
        $filteredIndices = @()
        $index = 0

        foreach ($device in $Devices) {
            if ($checkBox.Checked -or $device.DeviceName -ne "Unknown Device") {
                $rowIndex = $dataGridView.Rows.Add($device.DeviceName, $device.MSIStatus, $device.Action)
                $device | Add-Member -MemberType NoteProperty -Name "RowIndex" -Value $rowIndex -Force
                $filteredIndices += $index
            }
            $index++
        }

        return $filteredIndices
    }

    # Initial population of DataGridView
    $script:filteredIndices = Update-DataGridView

    # Checkbox event handler
    $checkBox.Add_CheckedChanged({
        $script:filteredIndices = Update-DataGridView
    })

    # Disable Unknown Devices button event handler
    $disableUnknownButton.Add_Click({
        $unknownCount = 0
        foreach ($device in $Devices) {
            if ($device.DeviceName -eq "Unknown Device") {
                $device.Action = "Disabled"
                $unknownCount++
            }
        }
        $script:filteredIndices = Update-DataGridView
        Write-Host "Set $unknownCount unknown devices to Disabled" -ForegroundColor Green
    })

    # Create OK and Cancel buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(600, 320)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.Text = "OK"
    $okButton.Add_Click({
        # Update device actions based on user selections
        for ($i = 0; $i -lt $dataGridView.Rows.Count; $i++) {
            $row = $dataGridView.Rows[$i]
            if ($i -lt $script:filteredIndices.Count) {
                $deviceIndex = $script:filteredIndices[$i]
                if ($null -ne $deviceIndex -and $null -ne $Devices[$deviceIndex] -and $null -ne $row.Cells["Action"].Value) {
                    $Devices[$deviceIndex].Action = $row.Cells["Action"].Value
                }
                else {
                    Write-Host "Warning: Skipping row $i due to invalid device index or Action value" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Warning: Row $i exceeds filtered indices count ($($script:filteredIndices.Count))" -ForegroundColor Yellow
            }
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(685, 320)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    # Add controls to the form
    $form.Controls.Add($checkBox)
    $form.Controls.Add($disableUnknownButton)
    $form.Controls.Add($dataGridView)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)

    # Show the form and return devices with updated actions
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $Devices
    }
    return @()
}

function Configure-SelectedDevices {
    param (
        [Parameter(Mandatory=$true)]
        [array]$SelectedDevices
    )

    $changesMade = 0

    Write-Host "`n===== Configuring MSI Settings for Devices =====" -ForegroundColor Cyan

    foreach ($device in $SelectedDevices) {
        if ($device.Action -eq $device.MSIStatus) {
            continue  # Skip if no change is needed
        }

        Write-Host "`nConfiguring: $($device.DeviceName)" -ForegroundColor Yellow
        Write-Host "Path: $($device.DevicePath)" -ForegroundColor White
        Write-Host "Action: $($device.Action)" -ForegroundColor White

        try {
            if ($device.Action -eq "Enabled") {
                New-ItemProperty -Path $device.DevicePath -Name "MSISupported" -Value 1 -PropertyType DWord -Force | Out-Null
                Write-Host "MSI Support: Enabled" -ForegroundColor Green
                $changesMade++
            }
            elseif ($device.Action -eq "Disabled") {
                New-ItemProperty -Path $device.DevicePath -Name "MSISupported" -Value 0 -PropertyType DWord -Force | Out-Null
                Write-Host "MSI Support: Disabled" -ForegroundColor Green
                $changesMade++
            }
            elseif ($device.Action -eq "Not Configured") {
                Remove-ItemProperty -Path $device.DevicePath -Name "MSISupported" -ErrorAction SilentlyContinue | Out-Null
                Write-Host "MSI Support: Removed (Not Configured)" -ForegroundColor Green
                $changesMade++
            }
        }
        catch {
            Write-Host "Configuration Failed: Unable to update MSI support" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor DarkRed
        }
    }

    return $changesMade
}

function Invoke-MSIConfiguration {
    # Verify administrative rights
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdministrator = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdministrator) {
        Write-Host "This script requires Administrator privileges!" -ForegroundColor Red
        Write-Host "Please run PowerShell as an Administrator." -ForegroundColor Yellow
        return
    }

    # Find devices
    $result = Find-MSIDevices
    $msiDevices = $result[0]
    $totalDevicesScanned = $result[1]

    if ($msiDevices.Count -eq 0) {
        Write-Host "`nNo devices found with MSI configuration options." -ForegroundColor Yellow
        return
    }

    # Show GUI for device configuration
    $configuredDevices = Show-DeviceSelectionGUI -Devices $msiDevices

    if ($configuredDevices.Count -eq 0) {
        Write-Host "`nNo changes made. Exiting." -ForegroundColor Yellow
        return
    }

    # Configure selected devices
    $changesMade = Configure-SelectedDevices -SelectedDevices $configuredDevices

    # Summary
    Write-Host "`n===== MSI Configuration Summary =====" -ForegroundColor Cyan
    Write-Host "Total Devices Scanned: $totalDevicesScanned" -ForegroundColor White
    Write-Host "Devices Configured: $changesMade" -ForegroundColor Green
}

# Call the main function
Invoke-MSIConfiguration