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
$computers | Select-Object Name | Export-Csv -Path $allComputersFile -NoTypeInformation
Write-Host "Exported all computers to $allComputersFile"

# Define the threshold for the last logon as three months ago
$threeMonthsAgo = (Get-Date).AddMonths(-3)

# Loop through each computer
foreach ($computer in $computers) {

    # Ping check using Test-Connection
    if (Test-Connection -ComputerName $computer.Name -Count 1 -Quiet) {
        try {
            # Use -ErrorAction Stop to catch the RPC error properly
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name -ErrorAction Stop
            $network = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $computer.Name -ErrorAction Stop | Where-Object { $_.IPEnabled }

            # Get the currently logged-in user
            $loggedInUser = if ($computerSystem.UserName) { $computerSystem.UserName } else { "No user logged in" }

            # Get network information
            $ipAddress = $network.IPAddress[0]
            $macAddress = $network.MACAddress
            $dnsHostName = $network.DNSHostName

            # Create a custom PowerShell object with the desired properties
            $result = [PSCustomObject]@{
                HostName     = $computer.Name
                IPAddress    = $ipAddress
                MACAddress   = $macAddress
                DNSHostName  = $dnsHostName
                LoggedInUser = $loggedInUser
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
            }
            Write-Warning $errorMessage
        }
    } else {
        # If ping fails, log the computer to the errorComputers array
        $pingFailMessage = "$($computer.Name) is unreachable, skipping..."
        $errorComputers += [PSCustomObject]@{
            ComputerName = $computer.Name
            Reason = "Unreachable (Ping failed)"
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