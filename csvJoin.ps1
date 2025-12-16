<#
.SYNOPSIS
    Parses Crestron device I/O join information from a .smw file into a .csv.

.DESCRIPTION
    This script parses input device (e.g., touch panel, XPanel) I/O join data from a specified .smw file.
    It correlates device joins (I/O) with Internal Signal Addresses (H), Signal Names (Nm), and normalizes 
    the sequential joins into I/A/S (Digital Input/Analog/Serial) and O/AO/SO (Digital Output/Analog Output/Serial Output) formats.
    This is highly useful for quickly referencing touch panel joins during user interface creation or for 
    documenting and standardizing control system interface joins.
    
    Logic decomposed and generated via iterative AI prompting and integrated by the author.

.PARAMETER InputPath (case sensitive)
    Specify the input .smw file PATH. If the script is in the same folder as the .smw than you can simply write the filename with the file extension. Example: program.smw

.PARAMETER ModelName (case sensitive)
    When prompted for the ModelName, you must provide the Nm=$ModelName of the correct device you want to parse the I/O joins from.
    Example: If the device is a touch panel, valid devices include TSW-560, TSW-570, TSW-770, etc.
    The correct ModelName can be found in the .smw file, typically near ObjVer=4 for touch panels, or use XPanel.

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

    [Parameter(Mandatory=$true)]
    [string]$ModelName,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath
)

# --- Create output path
$formatedInputPath = $InputPath -replace '\.[^.]*$', ''

if (-not $OutputPath) {
    $OutputPath = "$formatedInputPath"+"_"+"$ModelName"+"_signal_map.csv"
}

# --- Step 1: Read File Content & Build Global Signal Map (H -> Nm) ---
Write-Host "Reading file and building global signal map..."
try {
    $fileContent = Get-Content -Path $InputPath -Raw
}
catch {
    Write-Error "Error reading file '$InputPath': $($_.Exception.Message)"
    exit 1
}

$signalMap = @{}
$signalPattern = 'ObjTp=Sg\s+H=(\d+)\s+Nm=(.*?)(?:\s+SgTp=\d+)?\s*\]'
$signalMatches = [regex]::Matches($fileContent, $signalPattern, 'Singleline')

foreach ($match in $signalMatches) {
    $h = $match.Groups[1].Value
    $nm = $match.Groups[2].Value.Trim()
    if (-not $signalMap.ContainsKey($h)) {
        $signalMap[$h] = $nm
    }
}

# --- Step 2: Parse Device Block and Extract Input Signal Counts (Only I-Counts Required) ---
Write-Host "Searching for $ModelName block and extracting Input counts..."

$deviceBlockPattern = "(\[\s*ObjTp=Sm[^\]]*?Nm=$ModelName[^\]]*?ObjVer=[34][^\]]*?mI=\d+.*?\])"

$deviceBlockMatch = [regex]::Match($fileContent, $deviceBlockPattern, 'Singleline')

if (-not $deviceBlockMatch.Success) {
    Write-Error "Could not find $ModelName block in file. Check the ObjTp, Nm (Name), ObjVer, and mI attributes."
    exit 1
}

$deviceBlock = $deviceBlockMatch.Groups[1].Value

# Extract Input counts (I-joins)
# Using '\s+' boundary for robustness to ensure we match the full tag name (e.g., n1I) and not partial matches.
$n1I = if ([regex]::Match($deviceBlock, '\s+n1I=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n1I=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
$n2I = if ([regex]::Match($deviceBlock, '\s+n2I=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n2I=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
$mI  = if ([regex]::Match($deviceBlock, '\s+mI=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+mI=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }
$serial = $mI - ($n1I + $n2I)

# Extract Output counts (O-joins)
$n1O = if ([regex]::Match($deviceBlock, '\s+n1O=(\d+)', 'IgnoreCase').Success) { [int][regex]::Match($deviceBlock, '\s+n1O=(\d+)', 'IgnoreCase').Groups[1].Value } else { 0 }

# --- Deriving Output Counts from Input Counts ---
$n2O = $n2I
$mO = $mI

Write-Host "$ModelName Input Counts: Digital=$n1I, Analog=$n2I, Serial=$serial Total=$mI"
Write-Host "$ModelName Output Counts (Derived): Digital=$n1O, Analog=$n2O, Total=$mO"

if ($mI -eq 0) {
    Write-Error "Failed to extract total input signal count (mI). Exiting."
    exit 1
}

# --- Step 3: Define Signal Type Ranges ---
# Input Joins (I, A, S)
$iDigitalEnd = $n1I
$iAnalogEnd = $n1I + $n2I
# Output Joins (O, AO, SO) - Uses the same counts
$oDigitalEnd = $n1O
$oAnalogEnd = $n1O + $n2O

# --- Step 4: Process Input Joins (I#) and Normalize (I, A, S) ---
Write-Host "Processing Input Joins..."
$joinPattern = 'I(\d+)=(\d+)'
$joinMatches = [regex]::Matches($deviceBlock, $joinPattern)
$results = @()

foreach ($joinMatch in $joinMatches) {
    $iJoin = $joinMatch.Groups[1].Value
    $hAddr = $joinMatch.Groups[2].Value
    $iJoinNum = [int]$iJoin
    
    $realJoinNum = $iJoinNum
    $joinPrefix = "I" 
    $signalType = "Unmapped"
    $direction = "Input"
    
    $signalType = switch ($iJoinNum) {
        { $_ -le $iDigitalEnd } { 
            # Digital Input (I): I1 to I{n1I}
            $realJoinNum = $iJoinNum
            $joinPrefix = "I" 
            "Digital" 
        }
        { ($_ -gt $iDigitalEnd) -and ($_ -le $iAnalogEnd) } { 
            # Analog Input (A): A1 to A{n2I}. Offset by $n1I.
            $realJoinNum = $iJoinNum - $iDigitalEnd
            $joinPrefix = "A"
            "Analog" 
        }
        { ($_ -gt $iAnalogEnd) -and ($_ -le $mI) } { 
            # Serial Input (S): S1 to S{mI - iAnalogEnd}. Offset by ($n1I + $n2I).
            $realJoinNum = $iJoinNum - $iAnalogEnd
            $joinPrefix = "S"
            "Serial" 
        }
        default { "Unmapped" }
    }

    $signalName = if ($signalMap.ContainsKey($hAddr)) { $signalMap[$hAddr] } else { "UNKNOWN_SIGNAL_H" }

    $results += [PSCustomObject]@{
        "Join_Direction" = $direction
        "Join_Number"    = $realJoinNum 
        "Signal_Type"    = $signalType                  
        "Signal_Name"    = $signalName
    }
}


# --- Step 5: Process Output Joins (O#) and Normalize (O, AO, SO) ---
Write-Host "Processing Output Joins (using Input counts for ranges)..."
$joinPattern = 'O(\d+)=(\d+)'
$joinMatches = [regex]::Matches($deviceBlock, $joinPattern)

foreach ($joinMatch in $joinMatches) {
    $oJoin = $joinMatch.Groups[1].Value
    $hAddr = $joinMatch.Groups[2].Value
    $oJoinNum = [int]$oJoin
    
    $realJoinNum = $oJoinNum
    $joinPrefix = "O" 
    $signalType = "Unmapped"
    $direction = "Output"
    
    # Use the derived Output ranges ($oDigitalEnd, $oAnalogEnd)
    $signalType = switch ($oJoinNum) {
        { $_ -le $oDigitalEnd } { 
            # Digital Output (O): O1 to O{n1O}
            $realJoinNum = $oJoinNum
            $joinPrefix = "O" 
            "Digital" 
        }
        { ($_ -gt $oDigitalEnd) -and ($_ -le $oAnalogEnd) } { 
            # Analog Output (AO): AO1 to AO{n2O}. Offset by $n1O.
            $realJoinNum = $oJoinNum - $oDigitalEnd
            $joinPrefix = "AO" 
            "Analog" 
        }
        { ($_ -gt $oAnalogEnd) -and ($_ -le $mO) } { 
            # Serial Output (SO): SO1 to SO{mO - oAnalogEnd}. Offset by ($n1O + $n2O).
            $realJoinNum = $oJoinNum - $oAnalogEnd
            $joinPrefix = "SO" 
            "Serial" 
        }
        default { "Unmapped" }
    }

    $signalName = if ($signalMap.ContainsKey($hAddr)) { $signalMap[$hAddr] } else { "UNKNOWN_SIGNAL_H" }

    $results += [PSCustomObject]@{
        "Join_Direction" = $direction
        "Join_Number"    = $realJoinNum 
        "Signal_Type"    = $signalType                  
        "Signal_Name"    = $signalName
    }
}

# --- Step 6: Export to CSV ---
if ($results.Count -gt 0) {
    # Combine results from Step 4 and 5
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Delimiter ',' -Encoding UTF8
    Write-Host "Successfully mapped and categorized $($results.Count) $ModelName signals (Input and Output)."
    Write-Host "Output saved to '$OutputPath'."
}
else {
    Write-Host "No signals extracted. Check the input file format and path."

}

