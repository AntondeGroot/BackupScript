param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigPath
)

function Get-Config($path) {
    if (!(Test-Path $path)) { throw "Config file not found: $path" }
    return Get-Content $path -Raw | ConvertFrom-Json
}

function New-DirectoryIfMissing($path) {
    if (!(Test-Path $path)) { New-Item -ItemType Directory -Path $path | Out-Null }
}

function Write-Log($logFile, $msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $logFile -Value $line
}

function Get-ZipContentStats($zipPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        # Directory entries end with '/' and have no content; only count real files
        $files  = @($zip.Entries | Where-Object { -not $_.FullName.EndsWith("/") })
        $newest = $files | Sort-Object { $_.LastWriteTime } -Descending | Select-Object -First 1

        return [PSCustomObject]@{
            FileCount      = $files.Count
            NewestFile     = if ($newest) { $newest.FullName } else { $null }
            NewestModified = if ($newest) { $newest.LastWriteTime.DateTime } else { $null }
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Format-Age($date) {
    $days = [int][Math]::Floor(((Get-Date) - $date).TotalDays)
    if ($days -lt 1)  { return "today" }
    if ($days -eq 1)  { return "1 day ago" }
    if ($days -lt 60) { return "$days days ago" }
    return "{0} months ago" -f [Math]::Floor($days / 30)
}

function Send-StatusEmail($emailCfg, $subject, $body) {
    if (!(Test-Path $emailCfg.credentialFile)) {
        throw "Credential file not found: $($emailCfg.credentialFile)"
    }
    $cred = Import-Clixml $emailCfg.credentialFile

    Send-MailMessage `
        -From $emailCfg.from `
        -To $emailCfg.to `
        -Subject $subject `
        -Body $body `
        -SmtpServer $emailCfg.smtpServer `
        -Port $emailCfg.smtpPort `
        -UseSsl:([bool]$emailCfg.useSsl) `
        -Credential $cred
}

# --------------------------
# Main
# --------------------------
$config = Get-Config $ConfigPath
Write-Host "Backup jobs found: $($config.backups.Count)"

if ($null -eq $config.backups -or $config.backups.Count -eq 0) {
    throw "No backups defined in config.backups"
}

$results = @()
$overallOk = $true

foreach ($backup in $config.backups) {
    $backupName   = $backup.name
    $sourceFolder = $backup.sourceFolder.Trim()
    $backupFolder = $backup.backupFolder.Trim()
    $logFolder    = Join-Path $backupFolder $config.logFolderName

    New-DirectoryIfMissing $backupFolder
    New-DirectoryIfMissing $logFolder

    $logFile = Join-Path $logFolder ("backup-" + $backupName + (Get-Date -Format "yyyy-MM-dd") + ".log")
    $zipPath = $null
    $sizeMB  = $null
    $stats   = $null

    try {
        Write-Log $logFile "Starting Notes backup..."
        if (!(Test-Path $sourceFolder)) { throw "Source folder does not exist: $sourceFolder" }
        
        $year  = (Get-Date).ToString("yy")
        $month = (Get-Date).ToString("MM")
        $day   = (Get-Date).ToString("dd")

        $zipName = "$backupName-$year-$month-$day.zip"
        $zipPath = Join-Path $backupFolder $zipName

        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force
            Write-Log $logFile "Removed existing zip: $zipPath"
        }

        Write-Log $logFile "Compressing '$sourceFolder' to '$zipPath'..."
        Compress-Archive -Path (Join-Path $sourceFolder "*") -DestinationPath $zipPath -CompressionLevel Optimal

        if (!(Test-Path $zipPath)) { throw "ZIP file not created: $zipPath" }

        $zipInfo = Get-Item $zipPath
        $sizeMB  = [Math]::Round($zipInfo.Length / 1MB, 2)

        $stats = Get-ZipContentStats $zipPath

        Write-Log $logFile "Backup OK. Zip size: ${sizeMB} MB, files: $($stats.FileCount), newest: $($stats.NewestFile) ($($stats.NewestModified))"

        $results += [PSCustomObject]@{
            Name = $backupName
            Ok   = $true
            Zip  = $zipPath
            SizeMB = $sizeMB
            Stats = $stats
            Log   = $logFile
            Error = $null
        }
    }
    catch {
        $overallOk = $false
        $err = $_.Exception.Message
        try { Write-Log $logFile "ERROR: $err" } catch {}

        $results += [PSCustomObject]@{
            Name = $backupName
            Ok   = $false
            Zip  = $zipPath
            SizeMB = $sizeMB
            Stats = $stats
            Log  = $logFile
            Error = $err
        }
    }
}

# ----- Send one summary email -----
try {
    $dateStr = Get-Date -Format "yyyy-MM-dd"
    $failed = @($results | Where-Object { -not $_.Ok })
    $ok = @($results | Where-Object { $_.Ok })

    $subject = if ($overallOk) {
        "Backups SUCCESSFUL ($dateStr) - $($ok.Count) jobs"
    } else {
        "Backups COMPLETED WITH FAILURES ($dateStr) - OK: $($ok.Count), Failed: $($failed.Count)!"
    }

    $bodyLines = @()
    $bodyLines += "Backup summary - $dateStr"
    $bodyLines += ""
    foreach ($r in $results) {
        if ($r.Ok) {
            $bodyLines += "   $($r.Name) - $($r.SizeMB) MB"
            $bodyLines += "   Files: $($r.Stats.FileCount)"
            if ($r.Stats.NewestModified) {
                $newestStr = $r.Stats.NewestModified.ToString("yyyy-MM-dd HH:mm")
                $bodyLines += "   Last modified: $newestStr ($(Format-Age $r.Stats.NewestModified)) - $($r.Stats.NewestFile)"
            } else {
                $bodyLines += "   Last modified: n/a (zip contains no files!)"
            }
            $bodyLines += "   Zip: $($r.Zip)"
            $bodyLines += "   Log: $($r.Log)"
        } else {
            $bodyLines += "   $($r.Name)"
            $bodyLines += "   Error: $($r.Error)"
            $bodyLines += "   Log: $($r.Log)"
        }
        $bodyLines += ""
    }

    $body = ($bodyLines -join "`r`n")
    Send-StatusEmail $config.email $subject $body
}
catch {
    # Don’t hide backup results just because email failed
    Write-Host "ERROR sending summary email: $($_.Exception.Message)"
}

if ($overallOk) {
    exit 0
} else {
    exit 1
}