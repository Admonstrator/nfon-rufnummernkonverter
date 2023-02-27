# Define the path to the input CSV file
$csv_path = "area_codes.csv"

# Define the path to the output CSV file
$output_path = "output.csv"

# Define the function that takes a row of data as input
function DetermineAreaCode([string]$phone_number) {
    # Remove any non-numeric characters from the input phone number
    $phone_number = $phone_number -replace "[^0-9]"

    # Remove the country code from the phone number if it exists
    if ($phone_number.StartsWith("49")) {
        $phone_number = $phone_number.Substring(2)
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

# Import the input CSV file using Import-Csv with a semicolon delimiter
$input_data = Import-Csv -Path $args[0] -Delimiter ';'

# Initialize a progress bar
Write-Progress -Activity "Formatting phone numbers" -Status "Starting" -PercentComplete 0

# Loop over the rows in the input data and format the phone number
$output_data = foreach ($row in $input_data) {
    $formatted_number = DetermineAreaCode($row.Telefonnummer)
    [PSCustomObject]@{
        Vorname = $row.Vorname
        Nachname = $row.Nachname
        Telefonnummer = $formatted_number
    }
    # Update the progress bar
    $progress = $input_data.IndexOf($row) / $input_data.Count * 100
    Write-Progress -Activity "Formatting phone numbers ..." -Status "Formatting phone numbers" -PercentComplete $progress
}

# Complete the progress bar
Write-Progress -Activity "Formatting phone numbers" -Status "Complete" -PercentComplete 100

# Export the output data to a CSV file with a semicolon delimiter and UTF-8 encoding
$output_data
