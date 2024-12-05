function Find-and-ConfigureMSIDevices {
    # Base registry path for PCI devices
    $baseKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\PCI"
    
    # Initialize collections
    $msiConfigurableDevices = @()
    $totalDevicesScanned = 0
    $msiConfiguredCount = 0

    Write-Host "===== Scanning PCI Devices for MSI Configuration =====" -ForegroundColor Cyan
    Write-Host "Analyzing devices for MSI support and configuration..." -ForegroundColor Green

    # Get all PCI devices
    $devices = Get-ChildItem -Path $baseKey -Recurse -ErrorAction SilentlyContinue

    foreach ($device in $devices) {
        $totalDevicesScanned++

        try {
            # Check if the MSISupported subkey already exists
            $msiSupportedPath = Join-Path -Path $device.PSPath -ChildPath "MSISupported"
            $existingMsiSupport = Get-ItemProperty -Path $device.PSPath -Name "MSISupported" -ErrorAction SilentlyContinue

            # Determine if device needs MSI configuration
            $needsMsiConfiguration = $null -eq $existingMsiSupport

            if ($needsMsiConfiguration) {
                # Prepare device information for potential MSI configuration
                $deviceInfo = [PSCustomObject]@{
                    DevicePath = $device.PSPath
                    NeedsConfiguration = $true
                }
                $msiConfigurableDevices += $deviceInfo

                # Detailed output for configurable devices
                Write-Host "`n[MSI Configurable Device]" -ForegroundColor Yellow
                Write-Host "Device Path: $($device.PSPath)" -ForegroundColor White
                
                try {
                    # Attempt to add MSI support
                    New-ItemProperty -Path $device.PSPath -Name "MSISupported" -Value 1 -PropertyType DWord -Force | Out-Null
                    $msiConfiguredCount++

                    Write-Host "MSI Support Added: Successfully configured" -ForegroundColor Green
                }
                catch {
                    Write-Host "Configuration Failed: Unable to add MSI support" -ForegroundColor Red
                    Write-Host "Error: $_" -ForegroundColor DarkRed
                }
            }
        }
        catch {
            Write-Host "Error processing device: $($device.PSPath)" -ForegroundColor Red
            Write-Host "Error Details: $_" -ForegroundColor DarkRed
        }
    }

    # Summary of MSI configuration
    Write-Host "`n===== MSI Configuration Summary =====" -ForegroundColor Cyan
    Write-Host "Total Devices Scanned: $totalDevicesScanned" -ForegroundColor White
    Write-Host "Devices Configured for MSI: $msiConfiguredCount" -ForegroundColor Green

    # Return list of configurable devices
    return $msiConfigurableDevices
}

# Execute the function with administrative privileges
function Invoke-MSIConfiguration {
    # Verify administrative rights
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdministrator = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdministrator) {
        Write-Host "This script requires Administrator privileges!" -ForegroundColor Red
        Write-Host "Please run PowerShell as an Administrator." -ForegroundColor Yellow
        return
    }

    # Run the configuration
    $configurableDevices = Find-and-ConfigureMSIDevices
}

# Call the main function
Invoke-MSIConfiguration
