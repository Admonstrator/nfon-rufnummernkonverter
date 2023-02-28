![Logo](readme/header.png)

# ConvertTo-NFONPhoneNumbers.ps1

Das Skript `ConvertTo-NFONPhoneNumbers.ps1` formatiert Telefonnummern in einer CSV-Datei gemäß dem deutschen Telefonnummernformat. Das Skript liest eine CSV-Eingabedatei mit einem Semikolon als Trennzeichen, formatiert die Telefonnummern entsprechend dem deutschen Format und schreibt die formatierten Daten in eine CSV-Ausgabedatei.

## Installation

1. Laden Sie das Skript herunter und speichern Sie es in einem beliebigen Verzeichnis auf Ihrem Computer.
2. Laden Sie die Datei `area_codes.csv` herunter und speichern Sie sie im Verzeichnis `data` im gleichen Verzeichnis wie das Skript.
3. Erstellen Sie eine Eingabedatei im CSV-Format mit den Spalten `Vorname,Nachname,Telefonnummer`. Die Datei muss ein Semikolon als Trennzeichen verwenden. Das Format der Telefonnummer ist nicht relevant, es wird automatisch in das deutsche Format konvertiert.

## Verwendung

```Powershell
.\ConvertTo-NFONPhoneNumbers.ps1 [-AreaCodeFile <string>] [-OutputFile <string>] [-FailFile <string>] [-InputFile] <string>
```

## Parameter

| Parameter      | Verwendung                                                                                                    |
| -------------- | ------------------------------------------------------------------------------------------------------------- |
| `AreaCodeFile` | Der Pfad zu der CSV-Datei, die die Vorwahlen für Deutschland enthält. Standardwert ist "data\area_codes.csv". |
| `OutputFile`   | Der Pfad zur Ausgabe-CSV-Datei. Der Standardwert ist "output\output.csv".                                     |
| `FailFile`     | Der Pfad zu der CSV-Datei für fehlgeschlagene Zeilen. Der Standardwert ist "output\fail.csv".                 |
| `InputFile`    | Der Pfad zur Eingabe-CSV-Datei. Dieser Parameter ist obligatorisch.                                           |

## Funktionsweise

1. Das Skript prüft, ob die CSV-Eingabedatei und die CSV-Datei mit den Vorwahlen existieren und ob sie Semikolon als Trennzeichen verwenden.
2. Das Skript liest die CSV-Eingabedatei und führt eine Schleife über jede Zeile aus, um die Telefonnummer entsprechend dem deutschen Format zu formatieren.
3. Die formatierten Daten werden in die Ausgabe-CSV-Datei geschrieben, und alle fehlerhaften Zeilen werden in die Fehler-CSV-Datei geschrieben.

## Beispiel

Um das Skript mit den Standardparametern auszuführen, öffnen Sie eine PowerShell-Konsole, navigieren Sie zu dem Verzeichnis, das das Skript enthält, und führen Sie den folgenden Befehl aus:

```powershell
.\ConvertTo-NFONPhoneNumbers.ps1 -InputFile input.csv
```

Dies liest die Datei `input.csv` im gleichen Verzeichnis, formatiert die Telefonnummern entsprechend dem deutschen Format und schreibt die formatierten Daten in die Datei `output.csv`.

## Einschränkungen

Das Script kann nur mit deutschen Telefonnummern umgehen, die mit einer Vorwahl beginnen. Es kann keine Telefonnummern ohne Vorwahl verarbeiten. Es kann auch keine Telefonnummern mit Ländervorwahl abseits von +49 oder 0049 verarbeiten.

## Abhängigkeiten

Für die Ausführung des Skripts ist PowerShell Version 5.1 oder höher erforderlich.

Das Skript benötigt außerdem eine CSV-Datei mit den Vorwahlen für Deutschland. Der Standard-Dateipfad ist `data\area_codes.csv`. Die Datei muss ein Semikolon als Trennzeichen enthalten.

----

# ConvertTo-NFONPhoneNumbers.ps1

The `ConvertTo-NFONPhoneNumbers.ps1` script formats phone numbers in a CSV file according to the German phone number format. The script reads an input CSV file with a semicolon delimiter, formats the phone numbers according to the German format, and writes the formatted data to an output CSV file.

## Installation

1. download the script and save it in any directory on your computer.
2. download the file `area_codes.csv` and save it in the directory `data` in the same directory as the script.
3. create an input file in CSV format with columns `Vorname,Nachname,Telefonnummer`. The file must use a semicolon as a separator. The format of the phone number is not relevant, it will be automatically converted to the German format.

## Usage

```powershell
.\Format-NFONPhoneNumbers.ps1 [-AreaCodeFile <string>] [-OutputFile <string>] [-FailFile <string>] [-InputFile] <string>
```

## Parameters

| Parameter      | Usage                                                                                                   |
| -------------- | ------------------------------------------------------------------------------------------------------- |
| `AreaCodeFile` | The path to the CSV file containing the area codes for Germany. Default value is "data\area_codes.csv". |
| `OutputFile`   | The path to the output CSV file. Default value is "output\output.csv".                                  |
| `FailFile`     | The path to the CSV file for failed rows. Default value is "output\fail.csv".                           |
| `InputFile`    | The path to the input CSV file. This parameter is mandatory.                                            |

## Functionality

1. The script checks if the input CSV file and the area codes CSV file exist, and if they use semicolon as delimiter.
2. The script reads the input CSV file and loops over each row to format the phone number according to the German format.
3. The formatted data is written to the output CSV file, and any failed rows are written to the fail CSV file.

## Example
To run the script with the default parameters, open a PowerShell console, navigate to the directory containing the script, and run the following command:

```powershell
.\ConvertTo-NFONPhoneNumbers.ps1 -InputFile input.csv
```

This will read the `input.csv` file in the same directory, format the phone numbers according to the German format, and write the formatted data to `output\output.csv` file.

## Restrictions

The script can only handle German phone numbers that start with an area code. It cannot handle phone numbers without area code. It also cannot handle phone numbers with country codes other than +49 or 0049.

## Dependencies

The script requires PowerShell version 5.1 or later to run.

The script also requires a CSV file containing the area codes for Germany. The default file path is `data\area_codes.csv`. The file must have semicolon as delimiter.