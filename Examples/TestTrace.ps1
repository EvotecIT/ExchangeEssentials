

$FilePathToCache = "C:\Temp\TraceCache1.xml"

if ($FilePathToCache -and (Test-Path -LiteralPath $FilePathToCache)) {
    $TraceCache = Import-Clixml -LiteralPath $FilePathToCache
} else {
    $TraceCache = [ordered] @{}
}

if ($TraceCache.LastDate) {
    $StartDate = Get-Date -Date $TraceCache.LastDate
    $EndDate = (Get-Date -Date $TraceCache.LastDate).AddHours(-1).AddMinutes(-15)
    Write-Color -Text 'Getting trace from ', $TraceCache.LastDate, ' to ', $EndDate -Color Green
    $Trace = Get-MessageTrace -StartDate $EndDate -EndDate $StartDate
    $TraceCache.FirstDate = ($Trace | Select-Object -First 1).Received
    $TraceCache.LastDate = ($Trace | Select-Object -Last 1).Received
    $TraceCache.Trace += $Trace
} else {
    $StartDate = (Get-Date)
    $EndDate = $StartDate.AddHours(-1)
    Write-Color -Text 'Getting trace from ', $StartDate, ' to ', $EndDate -Color Green
    $Trace = Get-MessageTrace
    $TraceCache.FirstDate = ($Trace | Select-Object -First 1).Received
    $TraceCache.LastDate = ($Trace | Select-Object -Last 1).Received
    $TraceCache.Trace += $Trace
}

$TraceCache | Export-Clixml -LiteralPath $FilePathToCache
$TraceCache.Trace.Count