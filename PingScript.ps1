# Import required module
Import-Module ImportExcel

# Prompt user to enter the Excel file path
$ExcelFile = "C:\TestFile\ServerTest.xlsx)"
$SheetName = "Sheet1"  # Change this if your sheet has a different name
$ServerColumn = "Servers"  # Change this to match your Server Name column
$IPColumn = "TestIP"  # Change this to match your IP Address column

# Read Server Name and IP Address from the Excel file
if (Test-Path $ExcelFile) {
    $Servers = Import-Excel -Path $ExcelFile -WorksheetName $SheetName | Select-Object $ServerColumn, $IPColumn
} else {
    Write-Host "File not found! Please check the path and try again." -ForegroundColor Red
    exit
}

# Function to ping servers
Function Test-Server {
    param ($ServerName, $IPAddress)
    $Ping = Test-Connection -ComputerName $IPAddress -Count 2 -Quiet -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        ServerName = $ServerName
        IPAddress = $IPAddress
        Status = if ($Ping) { "Online" } else { "Offline" }
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# Ping all servers and collect results
$Results = @()
foreach ($Entry in $Servers) {
    $Results += Test-Server -ServerName $Entry.$ServerColumn -IPAddress $Entry.$IPColumn
}

# Display results in a table
$Results | Format-Table -AutoSize

# (Optional) Save results to an Excel file
$OutputFile = "C:\Path\To\PingResults.xlsx"
$Results | Export-Excel -Path $OutputFile -WorksheetName "PingResults" -AutoSize -BoldTopRow

Write-Host "Ping results saved to: $OutputFile" -ForegroundColor Green
