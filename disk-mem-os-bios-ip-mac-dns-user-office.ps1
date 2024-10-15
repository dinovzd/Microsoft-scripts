# File paths for the output files
$availableComputersFile = "C:\Users\cratis\Desktop\available_computers.csv"
$errorComputersFile = "C:\Users\cratis\Desktop\error_computers.csv"
$allComputersFile = "C:\Users\cratis\Desktop\all_computers.csv"

# Initialize arrays to store valid and error computer objects
$validComputers = @()
$errorComputers = @()

# Gather the list of computers (using Get-ADComputer)
$computers = Get-ADComputer -Filter {OperatingSystem -notLike '*server*'} -Properties Name, OperatingSystem, OperatingSystemVersion, LastLogon, PasswordLastSet

# Export all computers to all_computers.csv
$computers | Select-Object Name, OperatingSystem, OperatingSystemVersion, PasswordLastSet | Export-Csv -Path $allComputersFile -NoTypeInformation
Write-Host "Exported all computers to $allComputersFile"

# Define the threshold for the last logon as three months ago
$threeMonthsAgo = (Get-Date).AddMonths(-3)

# Function to get the installed Office version
function Get-OfficeVersion {
    param (
        [string]$computerName
    )

    try {
        # Query the registry for Office installations
        $officeKey = Get-WmiObject -Class Win32_Product -ComputerName $computerName -ErrorAction Stop | 
            Where-Object { $_.Name -match "Microsoft Office" } | 
            Select-Object -First 1 -Property Name, Version

        if ($officeKey) {
            return "$($officeKey.Name) ($($officeKey.Version))"
        } else {
            return "Office not found"
        }
    } catch {
        return "Failed to retrieve Office information for $($computerName): $_"
    }
}

# Loop through each computer
foreach ($computer in $computers) {
    # Check PasswordLastSet date
    if ($computer.PasswordLastSet) {
        $passwordLastSetDate = $computer.PasswordLastSet
    } else {
        $passwordLastSetDate = $null
    }

    # Skip the computer if the password has not been set in the last three months
    if ($passwordLastSetDate -and $passwordLastSetDate -lt $threeMonthsAgo) {
        Write-Host "$($computer.Name) has not had a password set in the last three months, skipping..."
        $errorComputers += [PSCustomObject]@{
            ComputerName = $computer.Name
            Reason = "Password not set in the last 3 months"
            PasswordNotSetDate = $passwordLastSetDate
        }
        continue
    }

    # Ping check using Test-Connection
    if (Test-Connection -ComputerName $computer.Name -Count 1 -Quiet) {
        try {
            # Use -ErrorAction Stop to catch the RPC error properly
            $processor = Get-WmiObject -Class Win32_Processor -ComputerName $computer.Name -ErrorAction Stop
            $memory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $computer.Name -ErrorAction Stop
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name -ErrorAction Stop
            $bios = Get-WmiObject -Class Win32_BIOS -ComputerName $computer.Name -ErrorAction Stop
            $disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $computer.Name -ErrorAction Stop | Where-Object { $_.DriveType -eq 3 }
            $network = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $computer.Name -ErrorAction Stop | Where-Object { $_.IPEnabled }

            # Calculate total memory in GB
            $totalMemoryGB = [math]::round(($memory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)

            # Get the Office version
            $officeVersion = Get-OfficeVersion -computerName $computer.Name

            # Get the currently logged-in user
            $loggedInUser = if ($computerSystem.UserName) { $computerSystem.UserName } else { "No user logged in" }

            # Get additional network information
            $ipAddress = $network.IPAddress[0]
            $macAddress = $network.MACAddress
            $dnsHostName = $network.DNSHostName

            # Get additional system information
            $systemManufacturer = $computerSystem.Manufacturer
            $systemModel = $computerSystem.Model
            $biosVersion = $bios.SMBIOSBIOSVersion

            # Calculate total disk capacity
            $totalDiskCapacityGB = [math]::round(($disk | Measure-Object -Property Size -Sum).Sum / 1GB, 2)

            # Create a custom PowerShell object with all the desired properties
            $result = [PSCustomObject]@{
                SystemName            = $computer.Name
                ProcessorName         = $processor.Name
                MemoryCapacityGB      = $totalMemoryGB
                MemoryManufacturer    = $memory.Manufacturer | Select-Object -First 1
                OperatingSystem       = $computer.OperatingSystem
                OperatingSystemVersion = $computer.OperatingSystemVersion
                OfficeVersion         = $officeVersion
                LoggedInUser          = $loggedInUser
                PasswordLastSet       = $passwordLastSetDate
                SystemManufacturer    = $systemManufacturer
                SystemModel           = $systemModel
                BIOSVersion           = $biosVersion
                TotalDiskCapacityGB    = $totalDiskCapacityGB
                IPAddress             = $ipAddress
                MACAddress            = $macAddress
                DNSHostName           = $dnsHostName
            }

            # Add the result to the valid computers array
            $validComputers += $result

            # Also print the result to console
            $result | Format-Table -AutoSize
        } catch {
            # Check if the error message contains "The RPC server is unavailable"
            if ($_.Exception.Message -like "*RPC server is unavailable*") {
                $errorMessage = "RPC server is unavailable for $($computer.Name)"
            } else {
                # Output any other errors
                $errorMessage = "Failed to retrieve information for $($computer.Name): $_"
            }

            # Log errors to errorComputers array
            $errorComputers += [PSCustomObject]@{
                ComputerName = $computer.Name
                Reason = $errorMessage
                PasswordNotSetDate = $passwordLastSetDate
            }
            Write-Warning $errorMessage
        }
    } else {
        # If ping fails, log the computer to the errorComputers array
        $pingFailMessage = "$($computer.Name) is unreachable, skipping..."
        $errorComputers += [PSCustomObject]@{
            ComputerName = $computer.Name
            Reason = "Unreachable (Ping failed)"
            PasswordNotSetDate = $passwordLastSetDate
        }
        Write-Warning $pingFailMessage
    }
}

# Export the valid computers to a CSV file
if ($validComputers.Count -gt 0) {
    $validComputers | Export-Csv -Path $availableComputersFile -NoTypeInformation
    Write-Host "Exported valid computers to $availableComputersFile"
} else {
    Write-Host "No valid computers found to export."
}

# Export the error computers to a CSV file
if ($errorComputers.Count -gt 0) {
    $errorComputers | Export-Csv -Path $errorComputersFile -NoTypeInformation
    Write-Host "Exported error computers to $errorComputersFile"
} else {
    Write-Host "No error computers found to export."
}