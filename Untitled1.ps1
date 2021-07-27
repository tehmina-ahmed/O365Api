# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

Connect-ExchangeOnline -UserPrincipalName admin_wh@dev2healthcarecollaboration.onmicrosoft.com -DelegatedOrganization dev2healthcarecollaboration.onmicrosoft.com

#Modify the values for the following variables to configure the audit log search.
$logFile = "c:\AuditLogSearch\AuditLogSearchLog.txt"
$outputFile = "c:\AuditLogSearch\AuditLogRecords.json"
[DateTime]$start = [DateTime]::UtcNow.AddDays(-2)
[DateTime]$end = [DateTime]::UtcNow
$record = "AzureActiveDirectory"
$resultSize = 5000
$intervalMinutes = 720

#Start script
[DateTime]$currentStart = $start
[DateTime]$currentEnd = $start


Function Write-LogFile ([String]$Message) {
    $final = [DateTime]::Now.ToUniversalTime().ToString("s") + ":" + $Message
    $final | Out-File $logFile -Append
}

Write-LogFile "BEGIN: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$resultSize."
Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$resultSize"

$totalCount = 0
while ($true) {
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $end) {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd) {
        break
    }

    $sessionID = [Guid]::NewGuid().ToString() + "_" + "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    Write-LogFile "INFO: Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    $currentCount = 0

    $sw = [Diagnostics.StopWatch]::StartNew()
    do {
        
        $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize
        $results | Select-Object -Property AuditData | Export-csv -Path $outputFile -NoTypeInformation
        

        if (($results | Measure-Object).Count -ne 0) {
            $results | export-csv -Path $outputFile -Append -NoTypeInformation

            $currentTotal = $results[0].ResultCount
            $totalCount += $results.Count
            $currentCount += $results.Count
            Write-LogFile "INFO: Retrieved $($currentCount) audit records out of the total $($currentTotal)"

            if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
                $message = "INFO: Successfully retrieved $($currentTotal) audit records for the current time range. Moving on!"
                Write-LogFile $message
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                ""
                break
            }
        }
    }
    while (($results | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}

Write-LogFile "END: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$resultSize, total count: $totalCount."
Write-Host "Script complete! Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green


Add-AzureRmAccount
Get-AzureRmSubscription | select SubscriptionId
Set-AzureRmContext -SubscriptionID "a9392f3e-86c6-4a94-ae4f-4db4ce33e55a"
Get-AzureRmStorageAccount -ResourceGroupName "O365WebApi" -Name "o365storageaccount1" | Set-AzureStorageBlobContent –Container o365container -File C:\AuditLogSearch\AuditLogRecords.json
Write-Host "Uploaded to Blob Storage" -foregroundColor Green