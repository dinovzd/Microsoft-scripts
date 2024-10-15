# File paths for the output files
$availableComputersFile = "C:\Users\cratis\Desktop\available_computers.csv"
$errorFile = "C:\Users\cratis\Desktop\error.txt"
$allComputersFile = "C:\Users\cratis\Desktop\all_computers.csv"

# Clear the content of the error file if it already exists
Clear-Content -Path $errorFile -ErrorAction SilentlyContinue

# Initialize an empty array to store the valid computer objects
$validComputers = @()

# Gather the list of computers (using Get-ADComputer)
$computers = Get-ADComputer -Filter {OperatingSystem -notLike '*server*'} -Properties Name, OperatingSystem, OperatingSystemVersion, LastLogon

# Export all computers to all_computers.csv
$computers | Select-Object Name, OperatingSystem, OperatingSystemVersion | Export-Csv -Path $allComputersFile -NoTypeInformation
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
    # Convert last logon from filetime if it's available
    if ($computer.LastLogon) {
        $lastLogonDate = [DateTime]::FromFileTime($computer.LastLogon)
    } else {
        $lastLogonDate = $null
    }

    # Skip the computer if it has not logged on within the last three months
    if ($lastLogonDate -and $lastLogonDate -lt $threeMonthsAgo) {
        Write-Host "$($computer.Name) has not logged on in the last three months, skipping..."
        Add-Content -Path $errorFile -Value "$($computer.Name) - Not logged on in the last 3 months"
        continue
    }

    # Ping check using Test-Connection
    if (Test-Connection -ComputerName $computer.Name -Count 1 -Quiet) {
        try {
            # Use -ErrorAction Stop to catch the RPC error properly
            $processor = Get-WmiObject -Class Win32_Processor -ComputerName $computer.Name -ErrorAction Stop
            $memory = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $computer.Name -ErrorAction Stop
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name -ErrorAction Stop

            # Calculate total memory in GB
            $totalMemoryGB = [math]::round(($memory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)

            # Get the Office version
            $officeVersion = Get-OfficeVersion -computerName $computer.Name

            # Get the currently logged-in user
            $loggedInUser = if ($computerSystem.UserName) { $computerSystem.UserName } else { "No user logged in" }

            # Create a custom PowerShell object with the desired properties, including OS version, Office version, and logged-in user
            $result = [PSCustomObject]@{
                SystemName          = $processor.SystemName
                ProcessorName       = $processor.Name
                MemoryCapacityGB    = $totalMemoryGB
                MemoryManufacturer  = $memory.Manufacturer | Select-Object -First 1
                OperatingSystem     = $computer.OperatingSystem
                OperatingSystemVersion = $computer.OperatingSystemVersion
                OfficeVersion       = $officeVersion
                LoggedInUser        = $loggedInUser
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

            # Log errors to error.txt
            Add-Content -Path $errorFile -Value $errorMessage
            Write-Warning $errorMessage
        }
    } else {
        # If ping fails, display a warning and log it to error.txt
        $pingFailMessage = "$($computer.Name) is unreachable, skipping..."
        Add-Content -Path $errorFile -Value $pingFailMessage
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