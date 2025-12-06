<#
    Edit-QuickAccessPins.ps1
    - Shows current Quick Access folder entries
    - Lets you select which ones to unpin
#>

# Require Out-GridView (built into Windows PowerShell; for PowerShell 7, install Microsoft.PowerShell.GraphicalTools)
if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
    Write-Error "Out-GridView is not available. Run this in Windows PowerShell, or install the 'Microsoft.PowerShell.GraphicalTools' module for PowerShell 7."
    return
}

# CLSID for Quick Access / Home
$quickAccessNamespace = "shell:::{679f85cb-0220-4080-B29B-5540CC05AAB6}"

Write-Host "Reading Quick Access items..." -ForegroundColor Cyan

# Connect to Shell COM object
$shell = New-Object -ComObject shell.application

# All items in Quick Access (pinned + frequent)
try {
    $items = $shell.Namespace($quickAccessNamespace).Items()
} catch {
    Write-Error "Could not access Quick Access namespace. Are you running on Windows 10/11 with Quick Access enabled?"
    return
}

if (-not $items) {
    Write-Host "No Quick Access items found." -ForegroundColor Yellow
    return
}

# Build a list of folder entries, try to detect pinned via System.Home.IsPinned
$pinnedObjects = @()

foreach ($item in $items) {
    # Only interested in folders with a real path
    if (-not $item.IsFolder) { continue }
    $path = $item.Path
    if ([string]::IsNullOrWhiteSpace($path)) { continue }

    $isPinned = $null
    try {
        # This shell property exists and can indicate pinned state
        $isPinned = $item.ExtendedProperty("System.Home.IsPinned")
    } catch {
        # Ignore if unavailable; we'll just leave IsPinned as $null
    }

    $pinnedObjects += [pscustomobject]@{
        Name     = $item.Name
        Path     = $path
        IsPinned = if ($isPinned -ne $null) { [bool]$isPinned } else { $null }
        Item     = $item   # keep the underlying COM object so we can unpin later
    }
}

if (-not $pinnedObjects) {
    Write-Host "No Quick Access folder entries found." -ForegroundColor Yellow
    return
}

# Show selection UI
Write-Host "Launching selection window..." -ForegroundColor Cyan
Write-Host "Tip: Ctrl+click or Shift+click to select multiple rows, then click OK." -ForegroundColor DarkGray

$selection = $pinnedObjects `
    | Sort-Object Name `
    | Select-Object Name, Path, IsPinned, Item `
    | Out-GridView -Title "Select Quick Access items to REMOVE (then click OK)" -PassThru

if (-not $selection) {
    Write-Host "No items selected. Nothing changed." -ForegroundColor Yellow
    return
}

Write-Host ""
Write-Host "Unpinning selected items from Quick Access..." -ForegroundColor Cyan

foreach ($entry in $selection) {
    $item = $entry.Item
    $display = "$($entry.Name)  [$($entry.Path)]"

    # Some items use 'unpinfromhome', some 'removefromhome' (e.g. Music/Videos) :contentReference[oaicite:1]{index=1}
    $verbs = @("unpinfromhome", "removefromhome")

    $success = $false
    foreach ($verb in $verbs) {
        try {
            $item.InvokeVerb($verb)
            $success = $true
        } catch {
            # ignore; we'll try the next verb
        }
    }

    if ($success) {
        Write-Host "Removed: $display" -ForegroundColor Green
    } else {
        Write-Host "Failed to remove: $display" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Done. You may need to close and reopen File Explorer to see changes immediately." -ForegroundColor Cyan
