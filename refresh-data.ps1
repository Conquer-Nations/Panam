$ErrorActionPreference = 'Stop'

$src = 'C:\Users\LisaSandoval\Downloads\PAN AMERICAN WIRE MFG NOV 2025 - CUSTOMER.xlsx'
$out = Join-Path $PSScriptRoot 'data.json'

if (-not (Test-Path $src)) {
    Write-Host "ERROR: Source workbook not found at $src" -ForegroundColor Red
    Write-Host "Edit refresh-data.ps1 and update the `$src path if your file lives elsewhere." -ForegroundColor Yellow
    exit 1
}

Write-Host "Reading $src ..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
$wb = $xl.Workbooks.Open($src, 0, $true)
$xl.CalculateFullRebuild()

$excelEpoch = [DateTime]'1899-12-30'

function Convert-CellValue($cell) {
    $v = $cell.Value2
    if ($null -eq $v) { return $null }
    if ($v -is [string]) {
        $t = $v.Trim()
        if ($t -eq '') { return $null }
        return $t
    }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) {
        $fmt = $cell.NumberFormat
        if ($fmt -match 'yy|mm/|d|YYYY|MM/|DD') {
            try {
                $dt = $excelEpoch.AddDays([double]$v)
                if ($dt.Year -ge 1990 -and $dt.Year -le 2099) {
                    return $dt.ToString('yyyy-MM-dd')
                }
            } catch {}
        }
        return [double]$v
    }
    return "$v"
}

$result = [ordered]@{}
$result.exportedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm')
$result.sheets = [ordered]@{}

$includeSheets = @('Summary','Balance Sheet','INVOICE MONTHLY ','INVOICE WEEKLY ','10 G','11 G','14 G Updates')

foreach ($name in $includeSheets) {
    $ws = $wb.Worksheets.Item($name)
    $ur = $ws.UsedRange
    $rows = $ur.Rows.Count
    $cols = $ur.Columns.Count
    Write-Host ("  - " + $name + " (" + $rows + " x " + $cols + ")") -ForegroundColor Gray
    $sheetData = @{ rows = $rows; cols = $cols; cells = New-Object 'System.Collections.ArrayList' }
    if ($rows -eq 1 -and $cols -eq 1) {
        $sheetData.cells.Add(@(Convert-CellValue $ws.Cells.Item(1,1))) | Out-Null
    } else {
        for ($r = 1; $r -le $rows; $r++) {
            $rowOut = New-Object 'System.Collections.ArrayList'
            for ($c = 1; $c -le $cols; $c++) {
                $rowOut.Add((Convert-CellValue $ws.Cells.Item($r, $c))) | Out-Null
            }
            $sheetData.cells.Add($rowOut) | Out-Null
        }
    }
    $result.sheets[$name] = $sheetData
}

$wb.Close($false)
$xl.Quit() | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
[gc]::Collect()

$json = $result | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText($out, $json, [System.Text.UTF8Encoding]::new($false))

Write-Host ("`nWrote " + $out) -ForegroundColor Green
Write-Host ("Size: " + (Get-Item $out).Length + " bytes") -ForegroundColor Green
Write-Host "`nNow re-upload data.json to your host (Netlify Drop / GitHub / OneDrive) to publish the update." -ForegroundColor Yellow
