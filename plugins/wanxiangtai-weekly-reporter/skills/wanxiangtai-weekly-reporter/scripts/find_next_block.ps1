param(
    [Parameter(Mandatory = $true)][string]$SpreadsheetToken,
    [Parameter(Mandatory = $true)][string]$SheetId,
    [Parameter(Mandatory = $true)][string]$PeriodLabel,
    [int]$FirstBlockStart = 3,
    [int]$BlockHeight = 27,
    [int]$BlockStride = 28,
    [int]$ScanThroughRow = 5000
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)

function Invoke-LarkJson([string[]]$Arguments) {
    $raw = (& lark-cli @Arguments 2>&1) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "lark-cli failed: $raw" }
    return $raw | ConvertFrom-Json
}

function Get-Value($Rows, [int]$RowNumber, [string]$Column) {
    if (-not $Rows.ContainsKey($RowNumber)) { return '' }
    $property = $Rows[$RowNumber].psobject.Properties[$Column]
    if ($null -eq $property -or $null -eq $property.Value) { return '' }
    return [string]$property.Value
}

function Test-BlockEmpty($Rows, [int]$Start, [int]$Height) {
    for ($row = $Start; $row -lt $Start + $Height; $row++) {
        if (-not $Rows.ContainsKey($row)) { continue }
        foreach ($property in $Rows[$row].psobject.Properties) {
            if ($null -ne $property.Value -and [string]$property.Value -ne '') { return $false }
        }
    }
    return $true
}

$workbook = Invoke-LarkJson @(
    'sheets', '+workbook-info', '--as', 'user',
    '--spreadsheet-token', $SpreadsheetToken
)
$sheet = @($workbook.data.sheets | Where-Object { [string]$_.sheet_id -eq $SheetId })[0]
if ($null -eq $sheet) { throw "工作表不存在：$SheetId" }
$rowCount = [int]$sheet.grid_properties.row_count
$scanEnd = [math]::Min($ScanThroughRow, $rowCount)
if ($scanEnd -lt $FirstBlockStart) { throw "工作表行数不足：row_count=$rowCount" }

$response = Invoke-LarkJson @(
    'sheets', '+csv-get', '--as', 'user',
    '--spreadsheet-token', $SpreadsheetToken,
    '--sheet-id', $SheetId,
    '--range', "A${FirstBlockStart}:F${scanEnd}",
    '--rows-json'
)

$rows = @{}
foreach ($row in @($response.data.rows)) { $rows[[int]$row.row_number] = $row.values }
$periods = New-Object 'System.Collections.Generic.List[string]'
$actualEnd = $FirstBlockStart
if ([string]$response.data.actual_range -match '(\d+)$') { $actualEnd = [int]$Matches[1] }
$limit = [math]::Min($scanEnd, [math]::Max($FirstBlockStart + $BlockStride, $actualEnd + $BlockStride))

for ($start = $FirstBlockStart; $start -le $limit; $start += $BlockStride) {
    $title = Get-Value $rows $start 'A'
    $period = Get-Value $rows ($start + 1) 'F'
    if ($period -eq $PeriodLabel -or $title.Contains("统计周期：$PeriodLabel")) {
        throw "统计周期已存在，拒绝覆盖：$PeriodLabel / block_start=$start"
    }
    if ($period) { $periods.Add($period) }

    if ($start -eq $FirstBlockStart -and -not $period -and ($title -match '____|待填写|统计周期：\s*$')) {
        [ordered]@{
            ok = $true
            block_start = $start
            mode = 'reuse_first_template'
            existing_periods = @($periods)
        } | ConvertTo-Json -Depth 5
        exit 0
    }

    if (Test-BlockEmpty $rows $start $BlockHeight) {
        [ordered]@{
            ok = $true
            block_start = $start
            mode = 'initialize_append_block'
            existing_periods = @($periods)
        } | ConvertTo-Json -Depth 5
        exit 0
    }

    if (-not $period) {
        throw "发现没有周期标记的非空块，拒绝猜测：block_start=$start"
    }
}

throw "在扫描范围内未找到空白周期块：through_row=$scanEnd"
