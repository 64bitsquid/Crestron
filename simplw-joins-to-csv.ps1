<#
.SYNOPSIS
    Parses Crestron device I/O join information from a .smw file into a .csv.

.DESCRIPTION

    WARNING: Don't corrupt your .smw file. Make a backup and work off of that. I will not be held responsible as stated in the below agreement.
    
    This script parses input device (e.g., touch panel, XPanel) I/O join data from a specified .smw file.
    It correlates device joins (I/O) with Internal Signal Addresses (H), Signal Names (Nm), and normalizes 
    the sequential joins into I/A/S (Digital Input/Analog/Serial) and O/AO/SO (Digital Output/Analog Output/Serial Output) formats.
    This is highly useful for quickly referencing touch panel joins during user interface creation or for 
    documenting and standardizing control system interface joins.
    
    Logic decomposed and generated via iterative AI prompting and integrated by the author.

.PARAMETER InputPath
    Specify the input .smw file PATH. If the script is in the same folder as the .smw than you can simply write the filename with the file extension. Example: program.smw

.NOTES
    Author: Jason Griffiths
    Date: 2025-12-16
    
    ========================================================================================================
    Copyright (c) [2025] [Jason Griffiths]

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
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
    "XPanel",
    "TSW-560",
    "TSW-550",
    "TSW-570",
    "TSW-750",
    "TSW-760",
    "TSW-770",
    "TSW-1080"
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

# --- Step 1a: Build VTP Map (DvH -> TSAddr/IPID) ---
# This allows us to link the internal device ID (DvH) to the user-friendly IP ID.
Write-Host "Building VTP IP ID map..."
$vtpMap = @{}
# Regex searches for VTP blocks containing DvH and TSAddr
$vtpPattern = '\[\s*ObjTp=VTP[\s\S]*?DvH=(\d+)[\s\S]*?TSAddr=([0-9A-Fa-f]+)[\s\S]*?\]'
$vtpMatches = [regex]::Matches($fileContent, $vtpPattern, 'Singleline')

foreach ($m in $vtpMatches) {
    $dvh = $m.Groups[1].Value
    $ipId = $m.Groups[2].Value
    if (-not $vtpMap.ContainsKey($dvh)) {
        $vtpMap[$dvh] = $ipId
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
    # UPDATED: Now using Matches() to find ALL instances of the model, not just the first one.
    $deviceBlockPattern = "(\[\s*ObjTp=Sm[^\]]*?Nm=$CurrentModel[^\]]*?ObjVer=[234][^\]]*?mI=\d+.*?\])"
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

        # --- Extract IP ID Suffix ---
        $ipIdSuffix = ""
        $dvhMatch = [regex]::Match($deviceBlock, 'DvH=(\d+)', 'IgnoreCase')
        if ($dvhMatch.Success) {
            $dvhKey = $dvhMatch.Groups[1].Value
            if ($vtpMap.ContainsKey($dvhKey)) {
                $ipIdSuffix = "_" + $vtpMap[$dvhKey] # e.g., "_03" or "_1F"
            }
        }

        # Extract Input/Output counts using boundary fix
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
        # If user supplied OutputPath, we use it (WARN: Will overwrite if multiple panels exist)
        # Otherwise, we generate a name: Generic_TSW-770_03_map.csv
        $FinalOutput = ""
        if ($OutputPath) { 
            $FinalOutput = $OutputPath 
        } else {
            $FinalOutput = "$formatedInputPath" + "_" + "$CurrentModel" + "$ipIdSuffix" + "_map.csv"
        }

        if ($results.Count -gt 0) {
            $results | Export-Csv -Path $FinalOutput -NoTypeInformation -Delimiter ',' -Encoding UTF8
            Write-Host "  -> Exported: $FinalOutput" -ForegroundColor Green
        }
    }
}




