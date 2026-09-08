param(
    [Parameter(Mandatory = $true)][string]$SpreadsheetToken,
    [Parameter(Mandatory = $true)][string]$PlanJson,
    [Parameter(Mandatory = $true)][string]$EnrichmentJson,
    [string[]]$HistoryPlanJsons = @(),
    [Parameter(Mandatory = $true)][string]$SheetId,
    [Parameter(Mandatory = $true)][string]$SheetTitle,
    [Parameter(Mandatory = $true)][string]$NoteCol,
    [int]$BlockStart = 59
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Get-ColName([int]$Index) {
    $result = ''
    while ($Index -gt 0) {
        $remainder = ($Index - 1) % 26
        $result = ([char](65 + $remainder)) + $result
        $Index = [math]::Floor(($Index - 1) / 26)
    }
    return $result
}

function Normalize($Value) {
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace(',', '').Replace('%', '').Trim()
}

function Invoke-LarkJson([string[]]$Arguments) {
    $raw = (& lark-cli @Arguments 2>&1) -join "`n"
    if ($LASTEXITCODE -ne 0) { throw "lark-cli failed: $raw" }
    return $raw | ConvertFrom-Json
}

function Assert-Summary($Rows, [int]$PeriodRow, [int]$SummaryStart, $Expected, [string]$Label) {
    $periodValue = [string]$Rows[$PeriodRow].F
    if ($periodValue -ne [string]$Expected.period) { throw "$Label 周期不符：$periodValue != $($Expected.period)" }
    $expectedValues = @($Expected.summary.impressions,$Expected.summary.cr,$Expected.summary.ctr,$Expected.summary.directGmv,$Expected.summary.spend,$Expected.summary.totalRoi)
    for ($i=0; $i -lt 6; $i++) {
        $actual = Normalize $Rows[$SummaryStart + $i].F
        $wanted = Normalize $expectedValues[$i]
        if ($actual -ne $wanted) { throw "$Label 汇总不符 row=$($SummaryStart+$i)：$actual != $wanted" }
    }
}

function Inspect-Materials($Materials, [int]$CoverRow, [int]$TypeRow, [int]$DirectionRow, [int]$RoiRow, [int]$UrlRow) {
    if ($Materials.Count -eq 0) { return [pscustomobject]@{covers=0;types=0;directions=0;formulas=0;urls=0} }
    $last = Get-ColName (6 + $Materials.Count)
    $cover = Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"G${CoverRow}:${last}${CoverRow}",'--include','value')
    $coverCells = @($cover.data.ranges[0].cells[0])
    $coverCount = @($coverCells | Where-Object { $_.rich_text -and $_.rich_text[0].type -eq 'embed-image' }).Count
    $meta = Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"G${TypeRow}:${last}${DirectionRow}",'--include','value,data_validation')
    $metaRows = @($meta.data.ranges[0].cells)
    $typeCount = @($metaRows[0] | Where-Object { $_.value -eq '图片' -or $_.value -eq '视频' }).Count
    $directionCount = @($metaRows[1] | Where-Object { $_.multiple_values -and @($_.multiple_values).Count -gt 0 }).Count
    $formula = Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"G${RoiRow}:${last}${RoiRow}",'--include','formula')
    $formulaCount = @($formula.data.ranges[0].cells[0] | Where-Object { $_.formula }).Count
    $urlCount = 0
    if ($UrlRow -gt 0) {
        $urls = Invoke-LarkJson @('sheets','+csv-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"G${UrlRow}:${last}${UrlRow}",'--rows-json')
        if (@($urls.data.rows).Count -gt 0) { $urlCount = @($urls.data.rows[0].values.psobject.Properties.Value | Where-Object { [string]$_ -ne '' }).Count }
    }
    return [pscustomobject]@{covers=$coverCount;types=$typeCount;directions=$directionCount;formulas=$formulaCount;urls=$urlCount}
}

$plan = Get-Content -LiteralPath $PlanJson -Raw | ConvertFrom-Json
$enrichment = Get-Content -LiteralPath $EnrichmentJson -Raw | ConvertFrom-Json
$historyPlans = @($HistoryPlanJsons | ForEach-Object { Get-Content -LiteralPath $_ -Raw | ConvertFrom-Json })
$expectedHistoryCount = [math]::Floor(($BlockStart - 3) / 28)
if ($BlockStart -lt 3 -or (($BlockStart - 3) % 28) -ne 0) { throw "块起始行不符合模板步长：$BlockStart" }
if ($historyPlans.Count -ne $expectedHistoryCount) {
    throw "历史周期快照数量不完整：需要 $expectedHistoryCount，实际 $($historyPlans.Count)"
}
$images = @($plan.materials | Where-Object { $_.type -eq '图片' })
$videos = @($plan.materials | Where-Object { $_.type -eq '视频' })
$end = $BlockStart + 26

$values = Invoke-LarkJson @('sheets','+csv-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"A1:${NoteCol}${end}",'--rows-json')
$rows = @{}
foreach ($row in $values.data.rows) { $rows[[int]$row.row_number] = $row.values }

for ($i = 0; $i -lt $historyPlans.Count; $i++) {
    $historyStart = 3 + 28 * $i
    Assert-Summary $rows ($historyStart + 1) ($historyStart + 6) $historyPlans[$i] "历史周期 $($i + 1)"
}
if ([string]$rows[$BlockStart].A -ne "计划：${SheetTitle}｜统计周期：$($plan.period)") { throw '新周期标题不符' }
Assert-Summary $rows ($BlockStart + 1) ($BlockStart + 6) $plan '新周期图片区'
if ([string]$rows[$BlockStart + 14].F -ne [string]$plan.period) { throw '新周期视频区周期不符' }

$imageCounts = Inspect-Materials $images ($BlockStart+2) ($BlockStart+4) ($BlockStart+5) ($BlockStart+12) 0
$videoCounts = Inspect-Materials $videos ($BlockStart+15) ($BlockStart+18) ($BlockStart+19) ($BlockStart+26) ($BlockStart+17)
foreach ($field in @('covers','types','directions','formulas')) {
    if ([int]$imageCounts.$field -ne $images.Count) { throw "图片 $field 数量不符：$($imageCounts.$field)/$($images.Count)" }
    if ([int]$videoCounts.$field -ne $videos.Count) { throw "视频 $field 数量不符：$($videoCounts.$field)/$($videos.Count)" }
}
if ([int]$videoCounts.urls -ne $videos.Count) { throw "视频 URL 数量不符：$($videoCounts.urls)/$($videos.Count)" }

$bestColumns = @{}
if ($images.Count -gt 0) {
    $bestIndex = -1
    for ($i=0; $i -lt $images.Count; $i++) { if ([string]$images[$i].id -eq [string]$enrichment.best_image_id) { $bestIndex=$i; break } }
    if ($bestIndex -lt 0) { throw '最佳图片 ID 不在本期图片中' }
    $col = Get-ColName (7 + $bestIndex)
    $cell = (Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"${col}$($BlockStart+12)",'--include','formula,style')).data.ranges[0].cells[0][0]
    if ($cell.formula -notmatch "${col}$($BlockStart+9)" -or $cell.formula -notmatch "${col}$($BlockStart+10)") { throw '最佳图片 ROI 公式引用不符' }
    if ([string]$cell.cell_styles.background_color -notmatch '(?i)#FFF2CC') { throw "最佳图片未标黄：$($cell.cell_styles.background_color)" }
    $bestColumns.image = $col
}
if ($videos.Count -gt 0) {
    $bestIndex = -1
    for ($i=0; $i -lt $videos.Count; $i++) { if ([string]$videos[$i].id -eq [string]$enrichment.best_video_id) { $bestIndex=$i; break } }
    if ($bestIndex -lt 0) { throw '最佳视频 ID 不在本期视频中' }
    $col = Get-ColName (7 + $bestIndex)
    $cell = (Invoke-LarkJson @('sheets','+cells-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"${col}$($BlockStart+26)",'--include','formula,style')).data.ranges[0].cells[0][0]
    if ($cell.formula -notmatch "${col}$($BlockStart+23)" -or $cell.formula -notmatch "${col}$($BlockStart+24)") { throw '最佳视频 ROI 公式引用不符' }
    if ([string]$cell.cell_styles.background_color -notmatch '(?i)#FFF2CC') { throw "最佳视频未标黄：$($cell.cell_styles.background_color)" }
    $bestColumns.video = $col
}

[pscustomobject]@{sheet=$SheetTitle; history_periods=@($historyPlans|ForEach-Object{$_.period}); new_period=$plan.period; images=$images.Count; videos=$videos.Count; image_counts=$imageCounts; video_counts=$videoCounts; best_columns=$bestColumns; status='verified'} | ConvertTo-Json -Depth 5 -Compress
