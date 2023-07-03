<#
.SYNOPSIS
  This script reads an AGEFO export file and prepares it for the use with ConvertTo-NFONPhoneNumbers.ps1
.DESCRIPTION
  TThis script reads an AGEFO export file and prepares it for the use with ConvertTo-NFONPhoneNumbers.ps1
.PARAMETER OutputFile
  The CSV file where the prepared phone numbers will be written to.
.PARAMETER InputFile
  The CSV file containing the phone numbers to be formatted.
.INPUTS
  CSV file as an export from AGFEO TK-Suite.
.OUTPUTS
 None
.NOTES
  Version:        1.0
  Author:         Aaron Viehl
  Creation Date:  2023-07-01
.EXAMPLE
  .\PrepareFile-Agfeo.ps1 -InputFile "agfeo-import.csv"
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$OutputFile = "output\agfeo-import-ready.csv",
    [string]$InputFile = "agfeo-import.csv",
    [string]$FailedFile = "output\agfeo-import-failed.csv"
)
$ErrorActionPreference = "Stop"
$Version = "1.1"

#---------------------------------------------------------[Functions]--------------------------------------------------------
function CheckCSVFile($CSVFile) {
    # Check if the CSV file exists
    if (!(Test-Path $CSVFile)) {
        Log -Severity "Error" "The CSV file $CSVFile does not exist. The script cannot continue." -ForegroundColor Red
        break
    }
    else {
        #check if delimiter is semicolon
        $delimiter = (Get-Content $CSVFile | Select-Object -First 1).Split(";").Length
        if ($delimiter -eq 1) {
            Log -Severity "Error" "The CSV file $CSVFile does not use semicolon as delimiter. The script cannot continue." -ForegroundColor Red
            break
        }
    }
    Log -Severity "Info" "The CSV file $CSVFile looks good, continuing."
}

function Log {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Severity,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $date = Get-Date -Format "yyyy-MM-dd"
    $time = Get-Date -Format "HH:mm:ss"

    switch ($Severity) {
        "INFO" {
            $color = "Green"
        }
        "WARN" {
            $color = "Yellow"
        }
        "ERROR" {
            $color = "Red"
        }
    }
    $severityText = "[$Severity]"
    $severityText = "[$Severity]".ToUpper()
    $dateTimeText = "$date $time"
    $messageText = $Message

    $maxSeverityLength = 7
    $maxDateTimeLength = 19

    $formattedSeverity = $severityText.PadRight($maxSeverityLength)
    $formattedDateTime = $dateTimeText.PadRight($maxDateTimeLength)
    $formattedMessage = $messageText

    $logMessage = "$formattedSeverity $formattedDateTime $formattedMessage"
    Write-Host $logMessage -ForegroundColor $color
}

function Header {
    Log -Severity "Info" "NFON Rufnummernkonverter preparation for AGFEO by Aaron Viehl (Singleton Factory GmbH)"
    Log -Severity "Info" "Prepares an AGFEO export file for the use with ConvertTo-NFONPhoneNumbers.ps1"
    Log -Severity "Info" "Version: $Version"
    Log -Severity "Info" "======================="
}

#---------------------------------------------------------[Main]--------------------------------------------------------
# Prepare the output file
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}
if (!(Test-Path "output")) {
    New-Item -ItemType Directory -Path "output" | Out-Null
}

Header

# Import the CSV file
CheckCSVFile $InputFile
$csv = Import-Csv -Path $inputFile -Delimiter ';'

# Prepare progress bar
# Get number of rows in input file
$number_of_rows = $csv.Count

# Initialize a progress bar
Write-Progress -Activity "Preparing phone numbers" -Status "Starting" -PercentComplete 0

Log -Severity "Info" "Engine is heated up and ready to go!"
Log -Severity "Info" "Starting to process $number_of_rows rows of data ..."

# Create an array to hold the new data
$newData = @()
$failedData = @()

# Loop through each row in the CSV
foreach ($row in $csv) {
    # Create a new object for each phone number
    $phoneColumns = @{
        '[1] Nummer: Festnetz (geschäftlich)' = ' (Büro)';
        '[1] Nummer: Mobil (geschäftlich)'    = ' (Mobil)';
        '[2] Nummer: Festnetz (geschäftlich)' = ' (Büro2)';
        '[1] Nummer: Festnetz (privat)'       = ' (Privat)';
    }
    foreach ($phoneColumn in $phoneColumns.Keys) {
        if (![string]::IsNullOrEmpty($row.$phoneColumn)) {
            $obj = New-Object PSObject
            $failed = New-Object PSObject
            
            # Create a list of the fields
            $fields = @($row.'Kontakt: Firma', $row.'Kontakt: Vorname', $row.'Kontakt: Name')
            # Remove any empty fields
            $fields = $fields | Where-Object { ![string]::IsNullOrEmpty($_) }
            # Join the fields with a space
            $name = [string]::Join(' ', $fields)
            
            # If Name is >50 characters, add it to $failedNames
            if ($name.Length -gt 50) {
                $failed | Add-Member -MemberType NoteProperty -Name "Name" -Value ($name + $phoneColumns[$phoneColumn])
                $failed | Add-Member -MemberType NoteProperty -Name "Telefonnummer" -Value $row.$phoneColumn
                $failedData += $failed
            }
            else {
                # Add the properties to the object
                $obj | Add-Member -MemberType NoteProperty -Name "Name" -Value ($name + $phoneColumns[$phoneColumn])
                $obj | Add-Member -MemberType NoteProperty -Name "Telefonnummer" -Value $row.$phoneColumn
                # Add the new object to the array
                $newData += $obj
            }
        }
        # get current row number
        $current_row = $csv.IndexOf($row) + 1
        # Calculate the percentage of completion
        $percent_complete = ($current_row / $number_of_rows) * 100

        # Update the progress bar
        Write-Progress -Activity "Preparing phone numbers" -Status "Processing row $current_row of $number_of_rows" -PercentComplete $percent_complete
        Start-Sleep -Milliseconds 10
    }
}

# Export the new data to a CSV
$failedData | Export-Csv -Path $failedFile -NoTypeInformation -Delimiter ';'
$newData | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ';'

Log -Severity "INFO" "======================="
Log -Severity "INFO" "Input file: $InputFile"
Log -Severity "INFO" "Output file: $OutputFile"
Log -Severity "INFO" "======================="
Log -Severity "Info" "Total number of rows: $number_of_rows"
if ($failedData.Count -eq 0) {
    Log -Severity "Info" "All rows were processed successfully!"
}
else {
    Log -Severity "Error" "Some rows were not processed successfully!"
    Log -Severity "Error" "Total number of failed rows: $($failedData.Count)"
    Log -Severity "Error" "They failed because their name was too long (>50 chars)."
    Log -Severity "Error" "You can find them in $FailedFile"
}

Log -Severity "Info" "======================="
Log -Severity "Info" "You can now run .\ConvertTo-NFONPhoneNumbers.ps1 -InputFile $OutputFile"

# Ask if the user wants to run ConvertTo-NFONPhoneNumbers.ps1 now
$answer = Read-Host -Prompt "Do you want to run ConvertTo-NFONPhoneNumbers.ps1 now? (y/n)"
if ($answer -eq "y") {
    .\ConvertTo-NFONPhoneNumbers.ps1 -InputFile $OutputFile
}
