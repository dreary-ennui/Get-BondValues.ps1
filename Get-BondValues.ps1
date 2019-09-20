<#
    .SYNOPSIS
    This PowerShell script is intended to fetch U.S. Savings Bond values by using the web form located at https://www.treasurydirect.gov/BC/SBCPrice using an input CSV.

    .DESCRIPTION
    Calling the script from the command line, provide it a date of redemption, an input CSV, and a filename for an output CSV. Please note I take no responsibility for this script! If the IRS or whoever shows up...

    .EXAMPLE
    .\Get-BondValues.ps1 -redemptionDate "01/2009" -inputCSV ".\inputCSV_Example.csv" -outputCSV ".\outputCSV_Example.csv"

    .PARAMETER redemptionDate
    String. Provide the date of redemption in the format "mm/yyyy". Each bond's values will be calculated according to this date. This is the "Value as of:" field on the website.

    .PARAMETER inputCSV
    String. Provide a CSV with fields Serial, Series, Denomination, and DateIssued for each bond.

    .PARAMETER outputCSV
    String. Provide a destination filename for the CSV output.

    .INPUTS
    None

    .OUTPUTS
    None
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({
        if ($_ -match "\d\d\/\d\d\d\d"){
            $true
        }
        else {
            Throw "$_ is not in proper format MM/yyyy."
        }
    })]
    [string]$redemptionDate,
    [Parameter(Mandatory=$true)]
    [ValidateScript({
        if (Test-Path $_){
            $true
        }
        else{
            throw "InputCSV file not found."
        }
    })]
    [string]$inputCSV,
    [Parameter(Mandatory=$true)]
    [string]$outputCSV
)

# The primary function that makes the web calls. Returns a computed bond object to the main script.
function Get-BondValues{
    param(
        [string]$redemptionDate,
        [string]$series,
        [string]$denomination,
        [string]$serialnumber,
        [string]$dateIssued
    )
    
    $formFields = [ordered]@{
        "RedemptionDate" = $redemptionDate
        "Series" = $series
        "Denomination" = $denomination
        "SerialNumber" = $serialnumber
        "IssueDate" = $dateIssued
        "btnAdd.x" = "CALCULATE"

        # I am not sure I need most of the ones below but hey.
        "SerialNumList"=""
        "IssueDateList"=""
        "SeriesList"=""
        "DenominationList"=""
        "IssuePriceList"=""
        "InterestList"=""
        "YTDInterestList"=""
        "ValueList"=""
        "InterestRateList"=""
        "NextAccrualDateList"=""
        "MaturityDateList"=""
        "NoteList"=""
        "OldRedemptionDate"="782"
        "ViewPos" = "0"
        "ViewType" = "Partial"
        "Version" = "6" 
    }

    try {
        # Make the web call.
        $result = Invoke-WebRequest -Uri $uri -method Post -Body $formFields -ContentType "application/x-www-form-urlencoded" -Headers $header -WebSession $session

        # If you input valid information, the website will tell you the data in a table called bnddata. I'm going to parse that out.
        $table = @($result.ParsedHtml.getElementsByClassName("bnddata"))
        
        # If you input invalid information, the website will tell you what's wrong. I'll capture that and tell you what happened in your outputCSV in the "Note" column.
        $errors = @($result.ParsedHtml.getElementsByClassName("errormessage"))
    }
    catch {
        $returnVar = "The following error(s) have occurred:General error making web request or interpreting the HTML returned."
    }

    if ($errors.count -ge 1){
        $returnVar = $errors.textcontent
    }

    elseif ($table.count -ge 1){
        #region The following code stolen from Lee Holmes. Thanks Lee! https://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
    
        $rows = @($table.rows)

        foreach ($row in $rows){
            $cells = @($row.Cells)
            ## If we've found a table header, remember its titles
            if($cells[0].tagName -eq "TH")
            {
                $titles = @($cells | ForEach-Object { ("" + $_.InnerText).Trim() })
                #region I added the following regex to clean up the titles a little bit. In this instance the table headers sometimes had line breaks in them.
                $titles = $titles -replace '[\W]',''
                #endregion
                continue
            }
            ## If we haven't found any table headers, make up names "P1", "P2", etc.
            if(-not $titles)
            {
                $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
                #region I added the following regex to clean up the titles a little bit. In this instance the table headers sometimes had line breaks in them.
                $titles = $titles -replace '[\W]',''
                #endregion
            }
            ## Now go through the cells in the the row. For each, try to find the
            ## title that represents that column and create a hashtable mapping those
            ## titles to content
            $resultObject = [Ordered] @{}
            for($counter = 0; $counter -lt $cells.Count; $counter++)
            {
                $title = $titles[$counter]
                if(-not $title) { continue }
                $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
            }
            ## And finally cast that hashtable to a PSCustomObject
            [PSCustomObject]$returnVar = [PSCustomObject]$resultObject
        }
    #endregion
    }
    return $returnVar
}


# Load the CSV into the primary bonds object
$bonds = import-csv $inputCSV

# Validate that the imported CSV matches our expectations. 
$expectedHeaders = "DateIssued","Denomination","Serial","Series"
$csvHeaders = $bonds | get-member -MemberType NoteProperty | Select-Object -ExpandProperty name
foreach ($header in $csvHeaders){
    if ($header -notin $expectedHeaders){
        throw "Your inputCSV contains invalid headers. Please provide only Serial, Series, Denomination, and DateIssued for each bond."
    }
}


#region Establishes initial web session
$uri = "https://www.treasurydirect.gov/BC/SBCPrice"

$header = @{
    "Referer"="https://www.treasurydirect.gov/BC/SBCPrice"
    "Accept"="text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3"
    "Accept-Encoding"="gzip, deflate, br"
    "Accept-Language"="en-US,en;q=0.9"
    "Origin"="https://www.treasurydirect.gov"
    "Sec-Fetch-Mode"="navigate"
    "Sec-Fetch-Site"="same-origin"
    "Sec-Fetch-User"="?1"
    "Upgrade-Insecure-Requests"="1"
}

try {
    $webRequest = Invoke-WebRequest -uri $uri -SessionVariable session -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
}
catch {
    throw "Error initiating web request."
}
#endregion

# Here we loop through each bond in the CSV, make the request, parse the response, and then add the results back to the original bond object.
foreach ($bond in $bonds){
    
    $bond.dateissued = get-date $bond.dateissued -Format "MM/yyyy"

    $result = Get-BondValues -redemptionDate $redemptionDate -series $bond.series -denomination $bond.denomination -serialnumber $bond.serial -dateIssued $bond.dateissued

    # If you put in bad info, we'll put the error message in the "Note" field.
    if ($result -like "*error*"){
        $bond | Add-Member -MemberType NoteProperty -name "redemptionDate" -value "$redemptionDate"
        $bond | Add-Member -MemberType NoteProperty -name "nextAccrual" -value ""
        $bond | Add-Member -MemberType NoteProperty -name "finalMaturity" -value ""
        $bond | Add-Member -MemberType NoteProperty -name "issuePrice" -value ""
        $bond | Add-Member -MemberType NoteProperty -name "interestAccrued" -value ""
        $bond | Add-Member -MemberType NoteProperty -name "interestRate" -value ""
        $bond | Add-Member -MemberType NoteProperty -name "Value" -value ""
        $errorValue = $result -Replace "The following error\(s\) have occurred:",""
        $bond | Add-Member -MemberType NoteProperty -name "Note" -value "$errorValue"
    }

    # Otherwise, here's the data you wanted. If you have a real "Note", it will show here. Check the https://www.treasurydirect.gov/BC/SBCPrice website for Note descriptions.
    else {
        $bond | Add-Member -MemberType NoteProperty -name "redemptionDate" -value "$redemptionDate"
        $bond | Add-Member -MemberType NoteProperty -name "nextAccrual" -value "$($result.nextAccrual)"
        $bond | Add-Member -MemberType NoteProperty -name "finalMaturity" -value "$($result.finalMaturity)"
        $bond | Add-Member -MemberType NoteProperty -name "issuePrice" -value "$($result.issuePrice)"
        $bond | Add-Member -MemberType NoteProperty -name "interestAccrued" -value "$($result.Interest)"
        $bond | Add-Member -MemberType NoteProperty -name "interestRate" -value "$($result.interestRate)"
        $bond | Add-Member -MemberType NoteProperty -name "Value" -value "$($result.value)"
        $bond | Add-Member -MemberType NoteProperty -name "Note" -value "$($result.note)"
    }   
}

# Finally, we export the container object to the specified destination.
try {
    $bonds | export-csv $outputCSV -NoTypeInformation -Append
}
catch {
    Throw "Had a problem exporting the CSV, unfortunately."
}
