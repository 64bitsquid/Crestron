<#
.SYNOPSIS
    Parses Crestron device I/O join information from a .smw file into a .csv.

.DESCRIPTION
    WARNING: Don't corrupt your .smw file. Make a backup and work off of that.
    
    This script parses input device (e.g., touch panel, XPanel) I/O join data from a specified .smw file.
    It correlates device joins (I/O) with Internal Signal Addresses (H), Signal Names (Nm), and normalizes 
    the sequential joins into I/A/S and O/AO/SO formats.
    
    Logic decomposed and generated via iterative AI prompting and integrated by the author.

.PARAMETER InputPath
    Specify the input .smw file PATH.

.NOTES
    Author: Jason Griffiths
    Date: 2025-12-16

    UPDATES:
    - Replaced VTP mapping with Dv (Device) mapping for IP IDs.
    - Uses Ad= attribute to find IP IDs, which is more reliable than SGD/VTP lookups.
    - Enforces exact model name matching to avoid "Buttons" sub-device confusion.

    ========================================================================================================
    Copyright (c) [2025] [Jason Griffiths]
    [MIT License Terms...]
    ========================================================================================================
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,

    [Parameter(Mandatory=$false)]
    [string]$ModelName,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath
)

# --- Format Input Path (Remove extension for cleaner naming) ---
$formatedInputPath = $InputPath -replace '\.[^.]*$', ''

# --- Define supported model list ---
$SupportedModels = @(
    "XPanel 2.0 Smart Graphics",
    "TSW-560",
    "TSW-570",
    "TSW-760",
    "TSW-770",
    "TSW-1060",
    "TSW-1070",
    "TSW-1080",
    "DGE-100",
)

# --- Determine Models to Process ---
$ModelsToProcess = @()
if (-not [string]::IsNullOrWhiteSpace($ModelName)) {
    $ModelsToProcess = @($ModelName)
} else {
    Write-Host "No Model Name entered. Scanning file for supported blocks..." -ForegroundColor Cyan
    $ModelsToProcess = $SupportedModels
}

# --- Step 1: Read File Content ---
Write-Host "Reading file..."
try {
    $fileContent = Get-Content -Path $InputPath -Raw
}
catch {
    Write-Error "Error reading file '$InputPath': $($_.Exception.Message)"
    exit 1
}

# --- Step 1a: Build Address Map (SmH -> Ad/IPID) ---
# Maps the Symbol Handle to the Hex IP ID found in ObjTp=Dv blocks.
Write-Host "Building Device Address map..."
$addrMap = @{}
$dvPattern = '\[\s*ObjTp=Dv[\s\S]*?Ad=([0-9A-Fa-f]+)[\s\S]*?SmH=(\d+)[\s\S]*?\]'
$dvMatches = [regex]::Matches($fileContent, $dvPattern, 'Singleline')

foreach ($m in $dvMatches) {
    $addr = $m.Groups[1].Value
    $smh = $m.Groups[2].Value
    if (-not $addrMap.ContainsKey($smh)) {
        $addrMap[$smh] = $addr
    }
}

# --- Step 1b: Build Global Signal Map (H -> Nm) ---
Write-Host "Building global signal map..."
$signalMap = @{}
$signalPattern = 'ObjTp=Sg\s+H=(\d+)\s+Nm=(.*?)(?:\s+SgTp=\d+)?\s*\]'
$signalMatches = [regex]::Matches($fileContent, $signalPattern, 'Singleline')

foreach ($match in $signalMatches) {
    $h = $match.Groups[1].Value
    $nm = $match.Groups[2].Value.Trim()
    if (-not $signalMap.ContainsKey($h)) { $signalMap[$h] = $nm }
}

# --- Loop Through Targeted Models ---
foreach ($CurrentModel in $ModelsToProcess) {
    Write-Host "Checking for $CurrentModel..." -ForegroundColor Yellow

    # Step 2: Parse Device Blocks using the "Good Regex"
    $deviceBlockPattern = "(\[\s*ObjTp=Sm[^\]]*?Nm=$CurrentModel\s*[\r\n][^\]]*?ObjVer=[234][^\]]*?mI=\d+.*?\])"
    $allDeviceMatches = [regex]::Matches($fileContent, $deviceBlockPattern, 'Singleline')

    if ($allDeviceMatches.Count -eq 0) {
        if ($ModelsToProcess.Count -eq 1) {
            Write-Warning "Could not find $CurrentModel block in file."
        }
        continue
    }

    # --- Process Each Instance Found ---
    foreach ($deviceMatch in $allDeviceMatches) {
        $deviceBlock = $deviceMatch.Groups[1].Value

        # --- Extract IP ID Suffix using SmH lookup ---
        $ipIdSuffix = ""
        $hMatch = [regex]::Match($deviceBlock, 'H=(\d+)', 'IgnoreCase')
        if ($hMatch.Success) {
            $hKey = $hMatch.Groups[1].Value
            if ($addrMap.ContainsKey($hKey)) {
                $ipIdSuffix = "_" + $addrMap[$hKey] # e.g., "_1F"
            }
        }

        # Extract Input/Output counts
        $n1I = if ([regex]::Match($deviceBlock, '\s+n1I=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n1I=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
        $n2I = if ([regex]::Match($deviceBlock, '\s+n2I=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n2I=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
        $mI  = if ([regex]::Match($deviceBlock, '\s+mI=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+mI=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
        $n1O = if ([regex]::Match($deviceBlock, '\s+n1O=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n1O=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
        
        $serial = $mI - ($n1I + $n2I)
        $n2O = $n2I
        $mO = $mI

        # Step 3: Define Ranges
        $iDigitalEnd = $n1I
        $iAnalogEnd = $n1I + $n2I
        $oDigitalEnd = $n1O
        $oAnalogEnd = $n1O + $n2O

        $results = @()

        # Step 4: Process Input Joins
        $iMatches = [regex]::Matches($deviceBlock, 'I(\d+)=(\d+)')
        foreach ($m in $iMatches) {
            $num = [int]$m.Groups[1].Value
            $h = $m.Groups[2].Value
            $type = switch ($num) {
                { $_ -le $iDigitalEnd } { "Digital" }
                { $_ -gt $iDigitalEnd -and $_ -le $iAnalogEnd } { "Analog" }
                { $_ -gt $iAnalogEnd -and $_ -le $mI } { "Serial" }
                default { "Unmapped" }
            }
            $realNum = switch ($type) {
                "Digital" { $num }
                "Analog"  { $num - $iDigitalEnd }
                "Serial"  { $num - $iAnalogEnd }
            }
            $results += [PSCustomObject]@{ "Join_Direction"="Input"; "Join_Number"=$realNum; "Signal_Type"=$type; "Signal_Name"=$signalMap[$h] }
        }

        # Step 5: Process Output Joins
        $oMatches = [regex]::Matches($deviceBlock, 'O(\d+)=(\d+)')
        foreach ($m in $oMatches) {
            $num = [int]$m.Groups[1].Value
            $h = $m.Groups[2].Value
            $type = switch ($num) {
                { $_ -le $oDigitalEnd } { "Digital" }
                { $_ -gt $oDigitalEnd -and $_ -le $oAnalogEnd } { "Analog" }
                { $_ -gt $oAnalogEnd -and $_ -le $mO } { "Serial" }
                default { "Unmapped" }
            }
            $realNum = switch ($type) {
                "Digital" { $num }
                "Analog"  { $num - $oDigitalEnd }
                "Serial"  { $num - $oAnalogEnd }
            }
            $results += [PSCustomObject]@{ "Join_Direction"="Output"; "Join_Number"=$realNum; "Signal_Type"=$type; "Signal_Name"=$signalMap[$h] }
        }

        # Step 6: Export CSV
        $FinalOutput = if ($OutputPath) { $OutputPath } else { "$formatedInputPath" + "_" + "$CurrentModel" + "$ipIdSuffix" + "_map.csv" }

        if ($results.Count -gt 0) {
            $results | Export-Csv -Path $FinalOutput -NoTypeInformation -Delimiter ',' -Encoding UTF8
            Write-Host "  -> Exported: $FinalOutput" -ForegroundColor Green
        }
    }
}






