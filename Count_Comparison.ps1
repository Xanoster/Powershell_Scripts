Import-Module SqlServer

$environment1 = Read-Host "Enter the first environment (Prod, Dev, QA)"
$environment2 = Read-Host "Enter the second environment (Prod, Dev, QA)"


$basePath = $PSScriptRoot
$outputDir = Join-Path $basePath "Output"
$htmlOutputPath = Join-Path $outputDir "$environment1 vs $environment2.html"


if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

function Get-ConnectionString {
    param ([string]$environment)
    switch ($environment) {
        'Prod' { return "Server=USBOSAPSQL16;Database=dwstage;Trusted_Connection=True;TrustServerCertificate=True;" }
        'Dev' { return "Server=usbocmdsql01;Database=dwstage;Trusted_Connection=True;TrustServerCertificate=True;" }
        'QA' { return "Server=usbocmdsql02;Database=dwstage;Trusted_Connection=True;TrustServerCertificate=True;" }
        default { throw "Invalid environment specified." }
    }
}

function Get-SQLData {
    param (
        [string]$environment,
        [string]$queryFile
    )

    $connectionString = Get-ConnectionString -environment $environment
    $query = Get-Content $queryFile | Out-String


    try {
        Write-Host "Executing SQL query for $environment..."
        $results = Invoke-Sqlcmd -Query $query -ConnectionString $connectionString -QueryTimeout 1800 -ErrorAction Stop
        
        if ($results.Count -eq 0) {
            Write-Host "No data returned for $environment."
            return $null
        }

        Write-Host "Query for $environment completed successfully."
        return $results | Select-Object dataset, Count_Records, Count_portfolio, Count_AsofDate
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to execute SQL query for $environment - $errorMessage"
        return $null
    }
}

function Convert-DatasetName {
    param (
        [string]$datasetName,
        [string]$environment
    )
    
    # Normalize 
    switch ($environment) {
        "Prod" { $datasetName = $datasetName -replace "^GVA_", "" }
        "Dev" { $datasetName = $datasetName -replace "^GVA_", "" }
        "QA" { $datasetName = $datasetName -replace "^QA\s", "" }
        default { return $datasetName } 
    }

    # handling Portfolios dataset
    if ($datasetName -like "*Portfolios*") {
        $datasetName = $datasetName -replace "_TD_DY$", "" 
    }
    return $datasetName
}


$queryFile1 = "$basePath\Query\Query$environment1.txt"
$queryFile2 = "$basePath\Query\Query$environment2.txt"

# Fetch data for the first environment
$dataEnv1 = Get-SQLData -environment $environment1 -queryFile $queryFile1
if (-not $dataEnv1) {
    Write-Host "No data fetched for $environment1."
    exit
}

$dataEnv2 = Get-SQLData -environment $environment2 -queryFile $queryFile2
if (-not $dataEnv2) {
    Write-Host "No data fetched for $environment2."
    exit
}

$results = @()

$dataEnv2Lookup = @{}
foreach ($row in $dataEnv2) {
    $datasetName = Convert-DatasetName -datasetName $row.dataset -environment $environment2
    if ($datasetName) {
        $dataEnv2Lookup[$datasetName] = $row
    }
}

# Compare data between the two environments
foreach ($row in $dataEnv1) {
    $convertedDatasetName = Convert-DatasetName -datasetName $row.dataset -environment $environment1
    $data2 = $dataEnv2Lookup[$convertedDatasetName]

    $results += [PSCustomObject]@{
        dataset                            = $row.dataset
        "Count_of_Records_Today_Env1"      = $row.Count_Records
        "Portfolio_Count_Today_Env1"       = $row.Count_portfolio
        "Distinct_Period_Check_Today_Env1" = $row.Count_AsofDate
        "Count_of_Records_Today_Env2"      = if ($data2) { $data2.Count_Records } else { $null }
        "Portfolio_Count_Today_Env2"       = if ($data2) { $data2.Count_portfolio } else { $null }
        "Distinct_Period_Check_Today_Env2" = if ($data2) { $data2.Count_AsofDate } else { $null }
        Count_Comparison                   = if ($data2) { if ($row.Count_Records -eq $data2.Count_Records) { "TRUE" } else { "FALSE" } } else { "FALSE" }
        Portfolio_Comparison               = if ($data2) { if ($row.Count_portfolio -eq $data2.Count_portfolio) { "TRUE" } else { "FALSE" } } else { "FALSE" }
        Distinct_Period_Check_Comparison   = if ($data2) { if ($row.Count_AsofDate -eq $data2.Count_AsofDate) { "TRUE" } else { "FALSE" } } else { "FALSE" }
    }
}

# Generate HTML
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; zoom: 0.9; }
        table { border-collapse: collapse; width: 100%; margin: 0; padding: 0; table-layout: auto; word-wrap: break-word; }
        th, td { border: 1px solid black; padding: 4px; text-align: center; }
        th { font-weight: bold; }
        .header-row { background-color: grey; font-weight: bold; }
        .sub-header-row { background-color: #AFC7E7; font-weight: bold; }
        .dataset-column { background-color: #AFC7E7; font-weight: bold; width: 15%; }
        .dataset-name { background-color: #D3D3D3; text-align: left; }
        .comparison-column { width: 6%; }
        .right-align { text-align: right; }
        .center-align { text-align: center; }
        .error { background-color: red; color: white; }
        .highlight { background-color: RED; }
    </style>
</head>
<body>
<table>
    <tr class="header-row">
        <td colspan="10">$environment1 vs $environment2 Counts Comparison</td>
    </tr>
    <tr class="sub-header-row">
        <td rowspan="2" class="dataset-column">Pipe_Name</td>
        <td colspan="3">$environment1</td>
        <td colspan="3">$environment2</td>
        <td rowspan="2" class="comparison-column">Count Comparison</td>
        <td rowspan="2" class="comparison-column">Portfolio Comparison</td>
        <td rowspan="2" class="comparison-column">Distinct Period Check Comparison</td>
    </tr>
    <tr class="sub-header-row">
        <td>Count of Records Today</td>
        <td>Portfolio Count Today</td>
        <td>Distinct Period Check Today</td>
        <td>Count of Records Today</td>
        <td>Portfolio Count Today</td>
        <td>Distinct Period Check Today</td>
    </tr>
"@

# Data rows
foreach ($result in $results) {
    $htmlContent += "<tr>"
    $htmlContent += "<td class='dataset-name'>$($result.dataset)</td>"

    $countClassEnv2 = if ($result.Count_Comparison -eq "FALSE") { 'class="error highlight right-align"' } else { 'class="right-align"' }
    $portfolioClassEnv2 = if ($result.Portfolio_Comparison -eq "FALSE") { 'class="error highlight right-align"' } else { 'class="right-align"' }
    $distinctPeriodCheckClassEnv2 = if ($result.Distinct_Period_Check_Comparison -eq "FALSE") { 'class="error highlight right-align"' } else { 'class="right-align"' }

    $countComparisonClass = if ($result.Count_Comparison -eq "FALSE") { 'class="error center-align highlight"' } else { 'class="center-align"' }
    $portfolioComparisonClass = if ($result.Portfolio_Comparison -eq "FALSE") { 'class="error center-align highlight"' } else { 'class="center-align"' }
    $distinctPeriodCheckComparisonClass = if ($result.Distinct_Period_Check_Comparison -eq "FALSE") { 'class="error center-align highlight"' } else { 'class="center-align"' }

    $htmlContent += "<td class='right-align'>$($result.Count_of_Records_Today_Env1)</td>"
    $htmlContent += "<td class='right-align'>$($result.Portfolio_Count_Today_Env1)</td>"
    $htmlContent += "<td class='right-align'>$($result.Distinct_Period_Check_Today_Env1)</td>"
    
    $htmlContent += "<td $countClassEnv2>$($result.Count_of_Records_Today_Env2)</td>"
    $htmlContent += "<td $portfolioClassEnv2>$($result.Portfolio_Count_Today_Env2)</td>"
    $htmlContent += "<td $distinctPeriodCheckClassEnv2>$($result.Distinct_Period_Check_Today_Env2)</td>"

    $htmlContent += "<td $countComparisonClass>$($result.Count_Comparison)</td>"
    $htmlContent += "<td $portfolioComparisonClass>$($result.Portfolio_Comparison)</td>"
    $htmlContent += "<td $distinctPeriodCheckComparisonClass>$($result.Distinct_Period_Check_Comparison)</td>"

    $htmlContent += "</tr>"
}

$htmlContent += @"
</table>
</body>
</html>
"@

$htmlContent | Out-File -FilePath $htmlOutputPath -Encoding UTF8

Write-Host "HTML report generated successfully: $htmlOutputPath"


#mail 
Write-Host "$(Get-Date -f s) Sending mail for $environment1 vs $environment2 comparison"
$message = "Please find attached the daily comparison report for $environment1 vs $environment2."
$smtpServer = "Smtp.gmail.com"
$smtpPort = 25
$emailFrom = "techypankaj@gmail.com"
#$emailTo = "pankajsain@gmail.com"
$password = "" 
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Port = $smtpPort
$smtp.Credentials = New-Object System.Net.NetworkCredential($emailFrom, $password)
$smtp.EnableSsl = $false


$attachment = New-Object System.Net.Mail.Attachment($htmlOutputPath)
$mailMessage = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo, "$environment1 vs $environment2 Comparison Report", $message)
$mailMessage.Attachments.Add($attachment)

Write-Host "$(Get-Date -f s) Sending HTML report as attachment"
try {
    $smtp.Send($mailMessage)
    Write-Host "$(Get-Date -f s) Email sent successfully"
}
catch {
    Write-Host "Failed to send email. Error: $_"
}
