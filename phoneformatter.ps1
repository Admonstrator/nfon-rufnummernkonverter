[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, Position=0)]
    [string]$PhoneNumber
)

# Define the path to the input CSV file
$csv_path = "area_codes.csv"

# Remove any non-numeric characters from the input phone number
$phone_number = $PhoneNumber -replace "[^0-9]"

# Check if the input phone number is a mobile phone number (starts with "01" or "15")
if ($phone_number.StartsWith("01") -or $phone_number.StartsWith("15")) {
    $formatted_number = "+49 ($($phone_number.Substring(1,3))) $($phone_number.Substring(4))"
    Write-Output "The number $formatted_number is a mobile phone number"
    return
}

# Import the CSV file using Import-Csv with a semicolon delimiter
$area_codes = Import-Csv -Path $csv_path -Delimiter ';'

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
                Write-Output "The number $formatted_number corresponds to $($matching_row.Ortsnetzkennzahl) - $($matching_row.Ortsnetzname)"
                return
            }
        }
    }
}

# If no match was found, output an error message
Write-Output "No matching area code found for $phone_number"