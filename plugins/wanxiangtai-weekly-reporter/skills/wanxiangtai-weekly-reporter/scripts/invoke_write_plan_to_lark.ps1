param(
    [Parameter(Mandatory = $true)][string]$SpreadsheetToken,
    [Parameter(Mandatory = $true)][string]$PlanJson,
    [Parameter(Mandatory = $true)][string]$EnrichmentJson,
    [Parameter(Mandatory = $true)][string]$SheetId,
    [Parameter(Mandatory = $true)][string]$SheetTitle,
    [Parameter(Mandatory = $true)][string]$AssetDir,
    [int]$BlockStart = 3,
    [string]$NoteCol,
    [switch]$SkipMediaUploads
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$Invariant = [Globalization.CultureInfo]::InvariantCulture

function Get-ColName([int]$Index) {
    $result = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $result = ([char](65 + $remainder)) + $result
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $result
}

function Convert-ToNumber($Value, [bool]$Percent = $false) {
    if ($null -eq $Value -or [string]$Value -eq '' -or [string]$Value -eq '-') { return 0.0 }
    $text = ([string]$Value).Replace(',', '').Trim()
    if ($text.EndsWith('%')) { $text = $text.Substring(0, $text.Length - 1); $Percent = $true }
    $result = [double]::Parse($text, $Invariant)
    if ($Percent) { return $result / 100.0 }
    return $result
}

function Get-MaterialDate([string]$Name) {
    $m = [regex]::Match($Name, '(?<!\d)(\d{1,2}\.\d{1,2})(?!\d)')
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($Name, '(?<!\d)(\d{3,4})(?!\d)')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

function New-ValueCell($Value) { return @{value = $Value} }
function New-FormulaCell([string]$Formula) { return @{formula = $Formula} }
function New-MultiCell($Values) {
    $items = @()
    foreach ($value in @($Values)) { $items += @{value = [string]$value} }
    return @{multiple_values = $items}
}

function New-Matrix {
    return New-Object 'System.Collections.Generic.List[object]'
}

function Add-Row($Matrix, [object[]]$Row) {
    $Matrix.Add([object[]]$Row)
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

$plan = Get-Content -LiteralPath $PlanJson -Raw | ConvertFrom-Json
$enrichment = Get-Content -LiteralPath $EnrichmentJson -Raw | ConvertFrom-Json
$images = @($plan.materials | Where-Object { $_.type -eq '图片' })
$videos = @($plan.materials | Where-Object { $_.type -eq '视频' })
if ($plan.materials.Count -ne $images.Count + $videos.Count) { throw '素材类型统计不一致' }
if ($images.Count -gt 193 -or $videos.Count -gt 193) {
    throw '素材超过飞书单表物理容量（G 起最多 193 列）；不得截断，请创建素材续表'
}
foreach ($item in $plan.materials) {
    if ($null -eq $enrichment.directions.psobject.Properties[[string]$item.id]) { throw "缺少方向：$($item.id)" }
}

$titleRow = $BlockStart
$imagePeriodRow = $titleRow + 1
$imageCoverRow = $titleRow + 2
$imageTypeRow = $titleRow + 4
$imageDirectionRow = $titleRow + 5
$imageMetricStart = $titleRow + 6
$imageGmvRow = $titleRow + 9
$imageSpendRow = $titleRow + 10
$imageRoiRow = $titleRow + 12
$videoPeriodRow = $titleRow + 14
$videoCoverRow = $titleRow + 15
$videoUrlRow = $titleRow + 17
$videoTypeRow = $titleRow + 18
$videoDirectionRow = $titleRow + 19
$videoMetricStart = $titleRow + 20
$videoGmvRow = $titleRow + 23
$videoSpendRow = $titleRow + 24
$videoRoiRow = $titleRow + 26

$summaryValues = @(
    (Convert-ToNumber $plan.summary.impressions),
    (Convert-ToNumber $plan.summary.cr $true),
    (Convert-ToNumber $plan.summary.ctr $true),
    (Convert-ToNumber $plan.summary.directGmv),
    (Convert-ToNumber $plan.summary.spend),
    (Convert-ToNumber $plan.summary.totalRoi)
)

$operations = @()
function Add-Cells([string]$Range, $Matrix) {
    $script:operations += @{shortcut = '+cells-set'; input = @{sheet_id = $SheetId; range = $Range; cells = $Matrix}}
}

$matrix = New-Matrix
Add-Row $matrix @((New-ValueCell "计划：${SheetTitle}｜统计周期：$($plan.period)"))
Add-Cells "A${titleRow}" $matrix

$imageLast = Get-ColName (6 + $images.Count)
$row = @((New-ValueCell $plan.period))
foreach ($item in $images) { $row += New-ValueCell (Get-MaterialDate ([string]$item.name)) }
$matrix = New-Matrix; Add-Row $matrix $row
Add-Cells "F${imagePeriodRow}:${imageLast}${imagePeriodRow}" $matrix

if ($images.Count -gt 0) {
    $types = @(); $dirs = @()
    foreach ($item in $images) {
        $types += New-ValueCell '图片'
        $dirs += New-MultiCell $enrichment.directions.psobject.Properties[[string]$item.id].Value
    }
    $matrix = New-Matrix; Add-Row $matrix $types; Add-Row $matrix $dirs
    Add-Cells "G${imageTypeRow}:${imageLast}${imageDirectionRow}" $matrix

    $fields = @('impressions','cr','ctr','directGmv','spend','totalRoi')
    $percents = @($false,$true,$true,$false,$false,$false)
    $matrix = New-Matrix
    for ($f=0; $f -lt $fields.Count; $f++) {
        $row = @((New-ValueCell $summaryValues[$f]))
        foreach ($item in $images) { $row += New-ValueCell (Convert-ToNumber $item.($fields[$f]) $percents[$f]) }
        Add-Row $matrix $row
    }
    $row = @()
    for ($index=6; $index -lt 7 + $images.Count; $index++) {
        $col = Get-ColName $index
        $row += New-FormulaCell ('=IF(OR({0}{1}="",{0}{2}="",{0}{2}=0),"",{0}{1}/{0}{2})' -f $col,$imageGmvRow,$imageSpendRow)
    }
    Add-Row $matrix $row
    Add-Cells "F${imageMetricStart}:${imageLast}${imageRoiRow}" $matrix
}

$videoLast = Get-ColName (6 + $videos.Count)
$row = @((New-ValueCell $plan.period))
foreach ($item in $videos) { $row += New-ValueCell (Get-MaterialDate ([string]$item.name)) }
$matrix = New-Matrix; Add-Row $matrix $row
Add-Cells "F${videoPeriodRow}:${videoLast}${videoPeriodRow}" $matrix

if ($videos.Count -gt 0) {
    $urls = @(); $types = @(); $dirs = @()
    foreach ($item in $videos) {
        $urls += New-ValueCell ([string]$item.url)
        $types += New-ValueCell '视频'
        $dirs += New-MultiCell $enrichment.directions.psobject.Properties[[string]$item.id].Value
    }
    $matrix = New-Matrix; Add-Row $matrix $urls; Add-Row $matrix $types; Add-Row $matrix $dirs
    Add-Cells "G${videoUrlRow}:${videoLast}${videoDirectionRow}" $matrix

    $fields = @('impressions','cr','ctr','directGmv','spend','totalRoi')
    $percents = @($false,$true,$true,$false,$false,$false)
    $matrix = New-Matrix
    for ($f=0; $f -lt $fields.Count; $f++) {
        $row = @((New-ValueCell $summaryValues[$f]))
        foreach ($item in $videos) { $row += New-ValueCell (Convert-ToNumber $item.($fields[$f]) $percents[$f]) }
        Add-Row $matrix $row
    }
    $row = @()
    for ($index=6; $index -lt 7 + $videos.Count; $index++) {
        $col = Get-ColName $index
        $row += New-FormulaCell ('=IF(OR({0}{1}="",{0}{2}="",{0}{2}=0),"",{0}{1}/{0}{2})' -f $col,$videoGmvRow,$videoSpendRow)
    }
    Add-Row $matrix $row
    Add-Cells "F${videoMetricStart}:${videoLast}${videoRoiRow}" $matrix
}

$missingImages = @($images | Where-Object { -not (Get-MaterialDate ([string]$_.name)) })
$missingVideos = @($videos | Where-Object { -not (Get-MaterialDate ([string]$_.name)) })
if (-not $NoteCol) { $NoteCol = Get-ColName ([math]::Max(24, 7 + [math]::Max($images.Count, $videos.Count))) }
if ($missingImages.Count -gt 0) {
    $parts = @($missingImages | ForEach-Object { "$($_.name) / 素材ID $($_.id)" })
    $matrix = New-Matrix; Add-Row $matrix @((New-ValueCell ('素材日期待核对：' + ($parts -join '；'))))
    Add-Cells "${NoteCol}${imagePeriodRow}" $matrix
}
if ($missingVideos.Count -gt 0) {
    $parts = @($missingVideos | ForEach-Object { "$($_.name) / 素材ID $($_.id)" })
    $matrix = New-Matrix; Add-Row $matrix @((New-ValueCell ('素材日期待核对：' + ($parts -join '；'))))
    Add-Cells "${NoteCol}${videoPeriodRow}" $matrix
}

$operationsJson = $operations | ConvertTo-Json -Depth 20 -Compress
$dryRun = Invoke-LarkWithInput @('sheets', '+batch-update', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--operations', '-', '--dry-run') $operationsJson
if ($dryRun -notmatch '"tool_name"\s*:\s*"batch_update"') { throw "batch dry-run did not complete: $dryRun" }
$resultRaw = Invoke-LarkWithInput @('sheets', '+batch-update', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--operations', '-', '--yes') $operationsJson
$result = $resultRaw | ConvertFrom-Json
if (-not $result.ok) { throw $resultRaw }

if (-not $SkipMediaUploads) {
  Push-Location -LiteralPath $AssetDir
  try {
    $index = 7
    foreach ($item in $images) {
        $col = Get-ColName $index
        $upload = Invoke-LarkJson @('sheets', '+cells-set-image', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "${col}${imageCoverRow}", '--image', ".\$($item.id).png")
        if (-not $upload.ok) { throw "图片上传失败：$($item.id)" }
        $index++
    }
    $index = 7
    foreach ($item in $videos) {
        $col = Get-ColName $index
        $upload = Invoke-LarkJson @('sheets', '+cells-set-image', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "${col}${videoCoverRow}", '--image', ".\$($item.id)-frame.png")
        if (-not $upload.ok) { throw "视频帧上传失败：$($item.id)" }
        $index++
    }
  } finally {
    Pop-Location
  }
}

$highlightRanges = @()
if ($images.Count -gt 0) {
    $bestIndex = [array]::FindIndex([object[]]$images, [Predicate[object]]{ param($x) [string]$x.id -eq [string]$enrichment.best_image_id })
    if ($bestIndex -lt 0) { throw '最佳图片 ID 不在本期图片中' }
    $bestCol = Get-ColName (7 + $bestIndex)
    $highlightRanges += "${SheetTitle}!${bestCol}${imagePeriodRow}:${bestCol}${imageRoiRow}"
}
if ($videos.Count -gt 0) {
    $bestIndex = [array]::FindIndex([object[]]$videos, [Predicate[object]]{ param($x) [string]$x.id -eq [string]$enrichment.best_video_id })
    if ($bestIndex -lt 0) { throw '最佳视频 ID 不在本期视频中' }
    $bestCol = Get-ColName (7 + $bestIndex)
    $highlightRanges += "${SheetTitle}!${bestCol}${videoPeriodRow}:${bestCol}${videoRoiRow}"
}
if ($highlightRanges.Count -gt 0) {
    $rangesJson = ConvertTo-Json -InputObject ([object[]]$highlightRanges) -Compress
    $styled = Invoke-LarkJson @('sheets', '+cells-batch-set-style', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--ranges', $rangesJson, '--background-color', '#FFF2CC')
    if (-not $styled.ok) { throw '最佳素材标黄失败' }
}

$imageCount = 0; $videoCount = 0
if ($images.Count -gt 0) {
    $check = Invoke-LarkJson @('sheets', '+cells-get', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "G${imageCoverRow}:${imageLast}${imageCoverRow}", '--include', 'value')
    $imageCount = [int]$check.data.returned_cell_count
}
if ($videos.Count -gt 0) {
    $check = Invoke-LarkJson @('sheets', '+cells-get', '--as', 'user', '--spreadsheet-token', $SpreadsheetToken, '--sheet-id', $SheetId, '--range', "G${videoCoverRow}:${videoLast}${videoCoverRow}", '--include', 'value')
    $videoCount = [int]$check.data.returned_cell_count
}
if ($imageCount -ne $images.Count -or $videoCount -ne $videos.Count) { throw "素材回读数量不符：image=$imageCount/$($images.Count) video=$videoCount/$($videos.Count)" }

[pscustomobject]@{sheet = $SheetTitle; images = $imageCount; videos = $videoCount; best_image_id = $enrichment.best_image_id; best_video_id = $enrichment.best_video_id; status = 'written_and_verified'} | ConvertTo-Json -Compress
