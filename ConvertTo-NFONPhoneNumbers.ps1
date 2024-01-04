<#
.SYNOPSIS
  This script converts phone numbers to the format used by NFON for an easy import into the NFON contact portal.
.DESCRIPTION
  This script converts phone numbers to the format used by NFON for an easy import into the NFON contact portal.
.PARAMETER AreaCodeFile
  The CSV file containing the area codes.
.PARAMETER OutputFile
  The CSV file where the formatted phone numbers will be written to.
.PARAMETER FailFile
  The CSV file where the phone numbers that could not be formatted will be written to.
.PARAMETER InputFile
  The CSV file containing the phone numbers to be formatted.
.INPUTS
  CSV file containing the phone numbers to be formatted.
.OUTPUTS
  Statistics about the conversion process and the CSV file containing the formatted phone numbers.
.NOTES
  Version:        1.0
  Author:         Aaron Viehl
  Creation Date:  2023-02-28
.EXAMPLE
  .\ConvertTo-NFONPhoneNumbers.ps1 -InputFile "input.csv"
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$AreaCodeFile = "data\area_codes.csv",
    [string]$OutputFile = "output\output.csv",
    [string]$FailFile = "output\fail.csv",
    [string]$InputFile = "input.csv"
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

# Define the function that takes a row of data as input
function DetermineAreaCode([string]$phone_number) {
    # Replace + with 00
    $phone_number = $phone_number -replace "^\+", "00"

    # Remove any non-numeric characters from the input phone number
    $phone_number = $phone_number -replace "[^0-9]"

    # If the phone number starts with 0049, it's a German number
    if ($phone_number.StartsWith("0049")) {
        $phone_number = $phone_number.Substring(4)
    }
    # If the phone number starts with 00 followed by any other digits, it's not a German number
    elseif ($phone_number -match "^00\d{2}") {
        return $null
    }

    # Import the CSV file using Import-Csv with a semicolon delimiter
    $area_codes = Import-Csv -Path $AreaCodeFile -Delimiter ';'

    # Define the minimum and maximum lengths of a valid area code
    $min_length = 2
    $max_length = 5

    # Loop over all possible substrings of the input phone number
    for ($i = 0; $i -lt $phone_number.Length; $i++) {
        for ($j = $min_length; $j -le $max_length; $j++) {
            # If the current substring is the right length, try to match it to an area code
            if ($i + $j -le $phone_number.Length) {
                $candidate_code = $phone_number.Substring($i, $j)
                $matching_row = $area_codes | Where-Object { $_.Ortsnetzkennzahl -eq $candidate_code }
                if ($matching_row) {
                    $formatted_number = "+49 ($candidate_code) $($phone_number.Substring($i + $j))"
                    return $formatted_number
                }
            }
        }
    }

    # If no match was found, assume the first four digits are the area code for formatting
    if ($phone_number.Length -ge 7) {
        $formatted_number = "+49 ($($phone_number.Substring(3,4))) $($phone_number.Substring(7))"
        return $formatted_number
    }

    # If the phone number is too short to have an area code, return the original phone number
    return $phone_number
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
    Log -Severity "Info" "NFON Rufnummernkonverter by Aaron Viehl (Singleton Factory GmbH)"
    Log -Severity "Info" "Prepares a CSV file containing phone numbers for import into the NFON Cloud PBX"
    Log -Severity "Info" "Version: $Version"
    Log -Severity "Info" "======================="
}

#---------------------------------------------------------[Main]--------------------------------------------------------
# Check if the PowerShell version is at least 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Log -Severity "Error" "This script requires PowerShell 7 or higher. Please upgrade your PowerShell version and try again." -ForegroundColor Red
    Log -Severity "Error" "You can download PowerShell 7 from https://aka.ms/PSWindows"
    break
}
    
# Cleaning up the output files and creating directories if necessary
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}
if (Test-Path $FailFile) {
    Remove-Item $FailFile
}
if (!(Test-Path "output")) {
    New-Item -ItemType Directory -Path "output" | Out-Null
}

Header
# Check all CSV files
CheckCSVFile($InputFile)
CheckCSVFile($AreaCodeFile)

# Import the input CSV file using Import-Csv with a semicolon delimiter
$input_data = Import-Csv -Path $InputFile -Delimiter ';'

# Get number of rows in input file
$number_of_rows = $input_data.Count

# Initialize counters
$fail_counter = 0
$convert_counter = 0
$skip_counter = 0

# Initialize a progress bar
Write-Progress -Activity "Formatting phone numbers" -Status "Starting" -PercentComplete 0
Log -Severity "Info" "Engine is heated up and ready to go!"
Log -Severity "Info" "Starting to process $number_of_rows rows of data ..."

# Loop over the rows in the input data and format the phone number
$output_data = foreach ($row in $input_data) {
    if ($row.Name.StartsWith("#")) {
        $skip_counter = $skip_counter + 1
        continue
    }
    # Skip rows with empty phone numbers
    if ($row.Telefonnummer -eq "") {
        $skip_counter = $skip_counter + 1
        continue
    }
    # Format the phone number
    $formatted_number = DetermineAreaCode($row.Telefonnummer)
    # Skip numbers with a different country code
    if (!$formatted_number) {
        # Write the row to the fail CSV file
        $fail_counter = $fail_counter + 1
        $row | Export-Csv -Path $FailFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation -Append
        continue
    }
    [PSCustomObject]@{
        displayName = $row.Name
        destination = $formatted_number
        visibleFor = ""
        vpnTargetNumber = ""
        vpnProvider = ""
    }
    $convert_counter = $convert_counter + 1
    $progress = $input_data.IndexOf($row) / $input_data.Count * 100
    $current_row = $input_data.IndexOf($row)
    Write-Progress -Activity "Formatting phone numbers ($current_row / $number_of_rows)..." -Status "Formatting phone numbers" -PercentComplete $progress
}

Write-Progress -Activity "Formatting phone numbers" -Status "Complete" -PercentComplete 100

# Write the output data to the output CSV file
$output_data | Export-Csv -Path $OutputFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation
Log -Severity "INFO" "======================="
Log -Severity "INFO" "Input file: $InputFile"
Log -Severity "INFO" "Output file: $OutputFile"
Log -Severity "INFO" "Failed file: $FailFile"
Log -Severity "INFO" "======================="
Log -Severity "INFO" "Conversion complete."
Log -Severity "INFO" "Total number of rows: $number_of_rows"
Log -Severity "INFO" "Total number of converted rows: $convert_counter"
Log -Severity "INFO" "Total number of skipped rows: $skip_counter"
Log -Severity "INFO" "Number of failed rows: $fail_counter"
Log -Severity "INFO" "======================="
Log -Severity "INFO" "Thank you for using this script."




