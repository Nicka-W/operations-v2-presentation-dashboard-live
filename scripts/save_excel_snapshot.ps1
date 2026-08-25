<#
save_excel_snapshot.ps1

Copies a workbook that is currently open/locked in the user's live Excel
session, WITHOUT attaching to that live session (attaching risks
0x800AC472 busy errors and can interrupt the user's other open
spreadsheets -- this was the root cause of an earlier fault in the
original Operations dashboard).

Approach:
1. Launch a SEPARATE, hidden Excel COM instance (not the user's).
2. Open the workbook read-only in that isolated instance.
3. SaveCopyAs to the destination temp path.
4. Close and release the isolated instance cleanly.

Usage:
    .\save_excel_snapshot.ps1 -SourcePath "C:\...\Workbook.xlsx" -DestPath "C:\Temp\snapshot.xlsx"
#>

param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$DestPath
)

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ReadOnly open avoids taking any lock on the source file.
    $workbook = $excel.Workbooks.Open($SourcePath, $null, $true)

    $destDir = Split-Path -Parent $DestPath
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    $workbook.SaveCopyAs($DestPath)
    Write-Host "Snapshot written to $DestPath"
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
