# Get-BondValues.ps1
This PowerShell script is intended to fetch U.S. Savings Bond values by using the web form located at https://www.treasurydirect.gov/BC/SBCPrice

## DESCRIPTION
Calling the script from the command line, provide it a date of redemption, an input CSV, and a filename for an output CSV. Please note I take no responsibility for this script! If the IRS or whoever shows up...

### EXAMPLE
.\Get-BondValues.ps1 -redemptionDate "01/2009" -inputCSV '.\inputCSV_Example.csv' -outputCSV ".\outputCSV_Example.csv"

### PARAMETER redemptionDate
String. Provide the date of redemption in the format "mm/yyyy". Each bond's values will be calculated according to this date. This is the "Value as of:" field on the website.

### PARAMETER inputCSV
String. Provide a CSV with fields Serial, Series, Denomination, and DateIssued for each bond.

### PARAMETER outputCSV
String. Provide a destination filename for the CSV output.

### INPUTS
None

### OUTPUTS
None
