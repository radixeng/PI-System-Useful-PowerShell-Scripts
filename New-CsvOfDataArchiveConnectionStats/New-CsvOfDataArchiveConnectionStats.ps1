Import-Module OSIsoft.PowerShell

# User Defined Variables

$PI_DATA_ARCHIVE_NAME = "grd001p-app027.guardian.corp"
$OUTPUT_FILE_NAME = "$PI_DATA_ARCHIVE_NAME ConnectionStatistics $(Get-Date -Format "yyyy-MM-dd HHmmss").csv"

# Main  
$ScriptFilePath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path $ScriptFilePath
$OutputDir = Join-Path $ScriptDir "output"
$OutputFilePath = Join-Path $OutputDir $OUTPUT_FILE_NAME

$DataArchiveConnection = Connect-PIDataArchive -PIDataArchiveMachineName $PI_DATA_ARCHIVE_NAME
$ConnectionStatistics = Get-PIConnectionStatistics -Connection $DataArchiveConnection

$ConnectionStatisticsPsCustomObjects = [PSCustomObject[]]($ConnectionStatistics | ForEach-Object {
    $_.GetEnumerator() | ForEach-Object {$aggregator = [Collections.Specialized.OrderedDictionary]::new()} {
        if ($_.Value -ne $null) {
            $aggregator.Add($_.Name, $_.Value)
        }
    } {[PSCustomObject]$aggregator}
})

New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
$ConnectionStatisticsPsCustomObjects | Export-Csv -Path $OutputFilePath -NoTypeInformation

# Compress as zip file
$OutputZipPath = Join-Path $OutputDir "$([IO.Path]::GetFileNameWithoutExtension($OUTPUT_FILE_NAME)).zip"
Compress-Archive @($OutputFilePath) -DestinationPath $OutputZipPath
Remove-Item -Path $OutputFilePath
