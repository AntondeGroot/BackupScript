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

function Get-ZipFileCount($zipPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        # Directory entries end with '/' and have no content; only count real files
        return @($zip.Entries | Where-Object { -not $_.FullName.EndsWith("/") }).Count
    }
    finally {
        $zip.Dispose()
    }
}

function Get-BackupStats($sourceFolder, $zipPath) {
    # Timestamps stored in the zip are not reliable, so take the newest file from the source folder.
    # Skip .obsidian: Obsidian rewrites its workspace files whenever the app is open.
    $newest = Get-ChildItem -Path $sourceFolder -File -Recurse |
        Where-Object { $_.FullName -notmatch '[\\/]\.obsidian[\\/]' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    return [PSCustomObject]@{
        FileCount      = Get-ZipFileCount $zipPath
        NewestFile     = if ($newest) { $newest.FullName.Substring($sourceFolder.TrimEnd('\').Length + 1) } else { $null }
        NewestModified = if ($newest) { $newest.LastWriteTime } else { $null }
    }
}

function Format-Age($date) {
    $days = [int][Math]::Floor(((Get-Date) - $date).TotalDays)
    if ($days -lt 1)  { return "today" }
    if ($days -eq 1)  { return "1 day ago" }
    if ($days -lt 60) { return "$days days ago" }
    return "{0} months ago" -f [Math]::Floor($days / 30)
}

# --------------------------
# Summary email (HTML)
# Inline styles and layout tables only: many mail clients (Outlook, Gmail) drop <style> blocks.
# --------------------------
$EmailColors = @{
    Ok         = "#1a7f37"
    OkTint     = "#dafbe1"
    Failed     = "#cf222e"
    FailedTint = "#ffebe9"
    Text       = "#1f2328"
    Muted      = "#656d76"
    Border     = "#e3e6ea"
    Subtle     = "#f6f8fa"
    Page       = "#eef1f4"
}
$EmailFont      = "-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif"
$EmailMonoFont  = "Consolas,Menlo,'Courier New',monospace"
$EmailCardStyle = "background:#ffffff;border:1px solid $($EmailColors.Border);border-radius:10px;"

function ConvertTo-HtmlText($value) {
    return [System.Net.WebUtility]::HtmlEncode([string]$value)
}

function Format-Size($sizeMB) {
    if ($sizeMB -ge 1024) { return "{0:N2} GB" -f ($sizeMB / 1024) }
    return "{0:N1} MB" -f [double]$sizeMB
}

function Get-StatusTheme($ok) {
    if ($ok) { return @{ Color = $EmailColors.Ok; Tint = $EmailColors.OkTint; Label = "OK"; Icon = "&#10003;" } }
    return @{ Color = $EmailColors.Failed; Tint = $EmailColors.FailedTint; Label = "FAILED"; Icon = "!" }
}

function Format-StatTile($label, $valueHtml, $noteHtml) {
    $labelStyle = "font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:$($EmailColors.Muted);"
    $valueStyle = "font-size:16px;font-weight:600;color:$($EmailColors.Text);padding-top:3px;"
    $noteStyle  = "font-size:12px;color:$($EmailColors.Muted);padding-top:1px;"
    return "<td width=`"33%`" valign=`"top`" style=`"padding:12px 14px;background:$($EmailColors.Subtle);border-radius:8px;`">" +
           "<div style=`"$labelStyle`">$label</div><div style=`"$valueStyle`">$valueHtml</div><div style=`"$noteStyle`">$noteHtml</div></td>"
}

function Format-StatRow($tiles) {
    $spacer = "<td width=`"8`" style=`"width:8px;`"></td>"
    return "<table role=`"presentation`" width=`"100%`" cellpadding=`"0`" cellspacing=`"0`"><tr>$($tiles -join $spacer)</tr></table>"
}

function Format-Notice($textHtml) {
    $style = "margin-top:12px;padding:10px 14px;background:$($EmailColors.FailedTint);border-radius:8px;color:$($EmailColors.Failed);font-size:14px;"
    return "<div style=`"$style`">$textHtml</div>"
}

function Format-PathList($paths) {
    $labelStyle = "padding:2px 12px 2px 0;color:$($EmailColors.Muted);font-size:12px;vertical-align:top;"
    $pathStyle  = "padding:2px 0;font-family:$EmailMonoFont;font-size:12px;color:$($EmailColors.Text);word-break:break-all;"
    $rows = foreach ($label in $paths.Keys) {
        "<tr><td style=`"$labelStyle`">$label</td><td style=`"$pathStyle`">$(ConvertTo-HtmlText $paths[$label])</td></tr>"
    }
    return "<div style=`"margin-top:14px;padding-top:10px;border-top:1px solid $($EmailColors.Border);`">" +
           "<table role=`"presentation`" cellpadding=`"0`" cellspacing=`"0`">$($rows -join '')</table></div>"
}

function Format-LastChangeTile($stats) {
    if (-not $stats.NewestModified) { return Format-StatTile "Last change" "&ndash;" "no files" }
    $age = Format-Age $stats.NewestModified
    $age = $age.Substring(0, 1).ToUpper() + $age.Substring(1)
    return Format-StatTile "Last change" $age $stats.NewestModified.ToString("yyyy-MM-dd HH:mm")
}

function Format-JobBody($result) {
    if (-not $result.Ok) {
        return (Format-Notice (ConvertTo-HtmlText $result.Error)) + (Format-PathList ([ordered]@{ Log = $result.Log }))
    }
    $stats = $result.Stats
    $body = Format-StatRow @(
        (Format-StatTile "Size" (Format-Size $result.SizeMB) "zipped"),
        (Format-StatTile "Files" ("{0:N0}" -f $stats.FileCount) "in archive"),
        (Format-LastChangeTile $stats)
    )
    $paths = [ordered]@{}
    if ($stats.NewestModified) { $paths.Latest = $stats.NewestFile }
    else { $body += Format-Notice "Source folder contains no files!" }
    $paths.Zip = $result.Zip
    $paths.Log = $result.Log
    return $body + (Format-PathList $paths)
}

function Format-JobCard($result) {
    $theme = Get-StatusTheme $result.Ok
    $pillStyle = "display:inline-block;padding:3px 10px;border-radius:999px;background:$($theme.Tint);color:$($theme.Color);font-size:11px;font-weight:700;letter-spacing:.6px;"
    return @"
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$EmailCardStyle margin:0 0 16px;">
  <tr><td style="padding:18px 20px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:14px;"><tr>
      <td style="font-size:17px;font-weight:600;color:$($EmailColors.Text);">$(ConvertTo-HtmlText $result.Name)</td>
      <td align="right"><span style="$pillStyle">$($theme.Label)</span></td>
    </tr></table>
    $(Format-JobBody $result)
  </td></tr>
</table>
"@
}

function Format-TotalsRow($results) {
    $ok = @($results | Where-Object { $_.Ok })
    $totalMB = 0; $totalFiles = 0
    foreach ($r in $ok) { $totalMB += $r.SizeMB; $totalFiles += $r.Stats.FileCount }
    return Format-StatRow @(
        (Format-StatTile "Jobs" "$($ok.Count) / $(@($results).Count)" "succeeded"),
        (Format-StatTile "Total size" (Format-Size $totalMB) "zipped"),
        (Format-StatTile "Total files" ("{0:N0}" -f $totalFiles) "backed up")
    )
}

function Format-EmailHeader($results, $overallOk) {
    $theme    = Get-StatusTheme $overallOk
    $headline = if ($overallOk) { "All backups succeeded" } else { "Some backups failed" }
    $date     = (Get-Date).ToString("dddd d MMMM yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    $iconStyle = "width:44px;height:44px;border-radius:22px;background:$($theme.Color);color:#ffffff;font-size:22px;font-weight:700;line-height:44px;text-align:center;"
    return @"
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="$EmailCardStyle border-top:4px solid $($theme.Color);margin:0 0 16px;">
  <tr><td style="padding:22px 20px 20px;">
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin-bottom:18px;"><tr>
      <td width="44" valign="middle"><div style="$iconStyle">$($theme.Icon)</div></td>
      <td valign="middle" style="padding-left:14px;">
        <div style="font-size:21px;font-weight:700;color:$($EmailColors.Text);">$headline</div>
        <div style="font-size:13px;color:$($EmailColors.Muted);padding-top:2px;">$date &middot; $(ConvertTo-HtmlText $env:COMPUTERNAME)</div>
      </td>
    </tr></table>
    $(Format-TotalsRow $results)
  </td></tr>
</table>
"@
}

function Format-SummaryEmail($results, $overallOk) {
    $cards = ($results | ForEach-Object { Format-JobCard $_ }) -join "`n"
    return @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:$($EmailColors.Page);">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:$($EmailColors.Page);"><tr><td align="center" style="padding:28px 12px;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;font-family:$EmailFont;">
    <tr><td>$(Format-EmailHeader $results $overallOk)</td></tr>
    <tr><td>$cards</td></tr>
    <tr><td style="padding:4px 0 8px;font-size:12px;color:$($EmailColors.Muted);text-align:center;">Sent by backup.ps1</td></tr>
  </table>
</td></tr></table>
</body></html>
"@
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
        -BodyAsHtml `
        -Encoding UTF8 `
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

        $stats = Get-BackupStats $sourceFolder $zipPath

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

    $body = Format-SummaryEmail $results $overallOk
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