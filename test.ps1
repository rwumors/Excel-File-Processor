# Make sure the necessary modules are available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Define the URL for the POST request
$Url = "https://pcsupport.lenovo.com/ca/en/api/v4/upsell/redport/getIbaseInfo"

# Define headers
$Headers = @{
    "Content-Type" = "application/json"
}

# Ask for the file path
$FilePath = Read-Host "Please provide the path to your CSV or Excel file"

# Check the file extension to determine how to read the file
$FileExt = [System.IO.Path]::GetExtension($FilePath)

if ($FileExt -eq ".csv") {
    # Import from CSV
    $SerialNumbers = Import-Csv -Path $FilePath
} elseif ($FileExt -eq ".xlsx") {
    # Import from Excel (using ImportExcel module)
    $SerialNumbers = Import-Excel -Path $FilePath
} else {
    Write-Host "Unsupported file type. Please use .csv or .xlsx."
    exit 1
}

# Create a list to store the results
$Results = @()

# Loop through each serial number in the file
foreach ($Row in $SerialNumbers) {
    $SerialNumber = $Row."Serial Number"  # Assumes the column header is "Serial Number"

    # Define the JSON body for the POST request
    $Body = @{
        serialNumber = $SerialNumber
    }

    # Convert the body to JSON format
    $BodyJson = $Body | ConvertTo-Json

    # Send the POST request
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method POST -Headers $Headers -Body $BodyJson
        $Data = $Response.data

        # Extract product name
        if ($Data.machineInfo.productName) {
            $ProductName = $Data.machineInfo.productName
        } else {
            $ProductName = "Product name not found"
        }

        # Add the result to the results array
        $Results += [PSCustomObject]@{
            "Serial Number" = $SerialNumber
            "Product Name"  = $ProductName
        }

        Write-Host "Processed Serial Number: ${SerialNumber} - Product Name: $ProductName"

    } catch {
        Write-Host "Error processing serial number ${SerialNumber}:"
        Write-Host $_.Exception.Message

        # Add the failed result to the results array
        $Results += [PSCustomObject]@{
            "Serial Number" = $SerialNumber
            "Product Name"  = "Error: Failed to retrieve data"
        }
    }
}

# Ask if the user wants to save the results to a file
$SaveResults = Read-Host "Do you want to save the results to a new CSV or Excel file? (yes/no)"

if ($SaveResults -eq "yes") {
    $OutputPath = Read-Host "Enter the path to save the results (e.g., results.csv or results.xlsx)"
    $OutputExt = [System.IO.Path]::GetExtension($OutputPath)

    if ($OutputExt -eq ".csv") {
        # Export to CSV
        $Results | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "Results saved to $OutputPath"
    } elseif ($OutputExt -eq ".xlsx") {
        # Export to Excel
        $Results | Export-Excel -Path $OutputPath
        Write-Host "Results saved to $OutputPath"
    } else {
        Write-Host "Unsupported output file type. Please use .csv or .xlsx."
    }
} else {
    Write-Host "Results were not saved to a file."
}

Write-Host "Processing complete."
