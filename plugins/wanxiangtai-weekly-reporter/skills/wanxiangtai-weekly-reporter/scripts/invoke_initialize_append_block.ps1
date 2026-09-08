param(
    [Parameter(Mandatory = $true)][string]$SpreadsheetToken,
    [Parameter(Mandatory = $true)][string]$SheetId,
    [Parameter(Mandatory = $true)][string]$SheetTitle,
    [Parameter(Mandatory = $true)][string]$LastCol,
    [int]$BlockStart = 31
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Get-ColIndex([string]$Name) {
    $result = 0
    foreach ($char in $Name.ToUpper().ToCharArray()) {
        $result = $result * 26 + ([int][char]$char - 64)
    }
    return $result
}

function Get-ColName([int]$Index) {
    $result = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $result = ([char](65 + $remainder)) + $result
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $result
}

function Invoke-LarkJson([string[]]$Arguments) {
    $raw = (& lark-cli @Arguments 2>&1) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "lark-cli failed: $raw" }
    return $raw | ConvertFrom-Json
}

function Invoke-LarkWithInput([string[]]$Arguments, [string]$InputText) {
    $raw = ($InputText | & lark-cli @Arguments 2>&1) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "lark-cli failed: $raw" }
    return $raw
}

$start = $BlockStart
$end = $start + 26
$lastColUpper = $LastCol.ToUpper()
$lastIndex = Get-ColIndex $lastColUpper
if ($lastIndex -lt 6) { throw '模板至少需要 A:F' }
$materialLast = Get-ColName ($lastIndex - 1)

$target = Invoke-LarkJson @('sheets', '+csv-get', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "A${start}:${lastColUpper}${end}", '--rows-json')
$nonempty = @()
foreach ($row in $target.data.rows) {
    $values = @($row.values.psobject.Properties.Value | Where-Object { $_ -ne $null -and [string]$_ -ne '' })
    if ($values.Count -gt 0) { $nonempty += $row.row_number }
}
if ($nonempty.Count -gt 0) { throw "目标追加区域非空，拒绝覆盖：rows=$($nonempty -join ',')" }

$layout = Invoke-LarkJson @('sheets', '+sheet-info', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "A${start}:${lastColUpper}${end}", '--include', 'merges')
if (@($layout.data.merged_cells).Count -gt 0) { throw '目标追加区域已有合并单元格，拒绝覆盖' }

$imagePeriod = $start + 1
$imageCoverStart = $start + 2
$imageCoverEnd = $start + 3
$imageEnd = $start + 12
$videoPeriod = $start + 14
$videoCoverStart = $start + 15
$videoCoverEnd = $start + 16
$videoEnd = $start + 26

$operations = @(
    @{shortcut = '+range-copy'; input = @{sheet_id = $SheetId; source_range = "A3:${lastColUpper}29"; target_range = "A${start}"; paste_type = 'formats'}},
    @{shortcut = '+range-copy'; input = @{sheet_id = $SheetId; source_range = 'A4:F29'; target_range = "A$($start + 1)"; paste_type = 'values'}},
    @{shortcut = '+cells-set'; input = @{sheet_id = $SheetId; range = "F${imageCoverStart}"; cells = ,@(@{value = '计划总计'})}},
    @{shortcut = '+cells-set'; input = @{sheet_id = $SheetId; range = "F${videoCoverStart}"; cells = ,@(@{value = '计划总计'})}},
    @{shortcut = '+cells-set-style'; input = @{sheet_id = $SheetId; range = "G${imagePeriod}:${materialLast}${imageEnd}"; background_color = '#FFFFFF'}},
    @{shortcut = '+cells-set-style'; input = @{sheet_id = $SheetId; range = "G${videoPeriod}:${materialLast}${videoEnd}"; background_color = '#FFFFFF'}}
)

$rowSizes = @(
    @("${start}", 36), @("$($start + 1)", 28), @("$($start + 2):$($start + 3)", 58),
    @("$($start + 4)", 28), @("$($start + 5)", 98), @("$($start + 6):$($start + 12)", 28),
    @("$($start + 13)", 12), @("$($start + 14)", 28), @("$($start + 15):$($start + 16)", 58),
    @("$($start + 17):$($start + 18)", 28), @("$($start + 19)", 106), @("$($start + 20):$($start + 26)", 28)
)
foreach ($entry in $rowSizes) {
    $operations += @{shortcut = '+rows-resize'; input = @{sheet_id = $SheetId; range = [string]$entry[0]; type = 'pixel'; size = [int]$entry[1]}}
}

$operationsJson = $operations | ConvertTo-Json -Depth 12 -Compress
$dryRun = Invoke-LarkWithInput @('sheets', '+batch-update', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--operations', '-', '--dry-run') $operationsJson
if ($dryRun -notmatch '"tool_name"\s*:\s*"batch_update"') { throw "batch dry-run did not complete: $dryRun" }
$resultRaw = Invoke-LarkWithInput @('sheets', '+batch-update', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--operations', '-', '--yes') $operationsJson
$result = $resultRaw | ConvertFrom-Json
if (-not $result.ok) { throw $resultRaw }

$verify = Invoke-LarkJson @('sheets', '+sheet-info', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "A${start}:${lastColUpper}${end}", '--include', 'merges,row_heights')
$expectedMerges = 2 * $lastIndex + 1
$actualMerges = @($verify.data.merged_cells).Count
if ($actualMerges -ne $expectedMerges) { throw "合并区验证失败：expected=$expectedMerges actual=$actualMerges" }

[pscustomobject]@{sheet = $SheetTitle; block = "A${start}:${lastColUpper}${end}"; merges = $actualMerges; status = 'initialized_and_verified'} | ConvertTo-Json -Compress
