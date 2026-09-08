param(
    [Parameter(Mandatory = $true)][string]$PlansFile,
    [Parameter(Mandatory = $true)][string]$WorkbookTitle,
    [string]$FolderToken,
    [ValidateRange(1, 193)][int]$MaterialCapacity = 24,
    [switch]$IncludeMonthlySheet,
    [switch]$ConfirmWrite
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Text.UTF8Encoding]::new($false)
$TemplateTitle = '_周报模板'
$DirectionOptions = @('用户信任','圈人群','正向','反向','功能','情感价值','使用感受','使用结果')

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

function Find-PropertyValue($Object, [string[]]$Names) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary] -or $Object -is [pscustomobject]) {
        foreach ($name in $Names) {
            $property = $Object.psobject.Properties[$name]
            if ($null -ne $property -and $property.Value) { return $property.Value }
        }
        foreach ($property in $Object.psobject.Properties) {
            $found = Find-PropertyValue $property.Value $Names
            if ($found) { return $found }
        }
    } elseif ($Object -is [System.Collections.IEnumerable] -and $Object -isnot [string]) {
        foreach ($item in $Object) {
            $found = Find-PropertyValue $item $Names
            if ($found) { return $found }
        }
    }
    return $null
}

function New-Matrix([int]$Rows, [int]$Cols) {
    $matrix = New-Object 'System.Collections.Generic.List[object]'
    for ($r = 0; $r -lt $Rows; $r++) {
        $row = New-Object 'System.Collections.Generic.List[object]'
        for ($c = 0; $c -lt $Cols; $c++) { $row.Add(@{value = ''}) }
        $matrix.Add([object[]]$row)
    }
    return $matrix
}

function Set-Cell($Matrix, [int]$Row, [int]$Col, $Value) {
    $Matrix[$Row - 1][$Col - 1] = @{value = $Value}
}

$plans = @(Get-Content -LiteralPath $PlansFile -Raw | ConvertFrom-Json)
if ($plans.Count -eq 0) { throw 'PlansFile 必须至少包含一个计划' }
$campaigns = @{}; $sheetNames = @{}
foreach ($plan in $plans) {
    foreach ($field in @('campaign_id','plan_name','sheet_name')) {
        if (-not [string]$plan.$field) { throw "计划缺少字段：$field" }
    }
    $campaign = [string]$plan.campaign_id
    $sheetName = [string]$plan.sheet_name
    if ($campaigns.ContainsKey($campaign)) { throw "campaign_id 重复：$campaign" }
    if ($sheetNames.ContainsKey($sheetName)) { throw "sheet_name 重复：$sheetName" }
    if ($sheetName -eq $TemplateTitle -or $sheetName -eq '计划月数据') { throw "工作表名为保留名称：$sheetName" }
    $campaigns[$campaign] = $true; $sheetNames[$sheetName] = $true
}

$materialFirstIndex = 7
$materialLastIndex = $materialFirstIndex + $MaterialCapacity - 1
$noteIndex = $materialLastIndex + 1
$lastCol = Get-ColName $noteIndex
$materialLastCol = Get-ColName $materialLastIndex
$matrix = New-Matrix 29 $noteIndex

Set-Cell $matrix 1 1 '展现量top：     ROI：1.3（以店铺花费投入，期待CTR和CR产出为准）'
Set-Cell $matrix 2 1 '筛选标准：roi1.3平均水平    目标：1.5    展现量：top3'
Set-Cell $matrix 3 1 '计划：________________｜统计周期：待填写'
Set-Cell $matrix 4 1 '图片'; Set-Cell $matrix 4 2 '填写本周图片测试方向'; Set-Cell $matrix 4 3 '数据截图'
Set-Cell $matrix 4 4 '结论'; Set-Cell $matrix 4 5 '投放日期'; Set-Cell $matrix 4 $noteIndex '填写结论 / 备注'
Set-Cell $matrix 5 5 '创意图 / 指标'; Set-Cell $matrix 5 6 '计划总计'
Set-Cell $matrix 7 5 '类型'; Set-Cell $matrix 8 5 '方向'
$metrics = @('展现','CR','CTR','直接成交金额','花费','总ROI','直接ROI')
for ($i = 0; $i -lt $metrics.Count; $i++) { Set-Cell $matrix (9 + $i) 5 $metrics[$i] }

Set-Cell $matrix 17 1 '视频'; Set-Cell $matrix 17 2 '填写本周视频测试方向'; Set-Cell $matrix 17 3 '数据截图'
Set-Cell $matrix 17 4 '结论'; Set-Cell $matrix 17 5 '投放日期'
Set-Cell $matrix 18 5 '创意视频 / 指标'; Set-Cell $matrix 18 6 '计划总计'
Set-Cell $matrix 20 5 '视频地址'; Set-Cell $matrix 21 5 '类型'; Set-Cell $matrix 22 5 '方向'
for ($i = 0; $i -lt $metrics.Count; $i++) { Set-Cell $matrix (23 + $i) 5 $metrics[$i] }

for ($i = 0; $i -lt $MaterialCapacity; $i++) {
    Set-Cell $matrix 5 ($materialFirstIndex + $i) ('素材{0:D2}（粘贴图片）' -f ($i + 1))
    Set-Cell $matrix 18 ($materialFirstIndex + $i) ('素材{0:D2}（粘贴视频帧）' -f ($i + 1))
}

$operations = New-Object 'System.Collections.Generic.List[object]'
if ($noteIndex -gt 20) {
    $operations.Add(@{shortcut='+dim-insert';input=@{sheet_id='__SHEET__';position='T';count=($noteIndex-20);inherit_style='none'}})
}
$operations.Add(@{shortcut='+cells-set';input=@{sheet_id='__SHEET__';range="A1:${lastCol}29";cells=$matrix;allow_overwrite=$false}})
foreach ($range in @("A1:${lastCol}1","A2:${lastCol}2","A3:${lastCol}3","F5:F6","F18:F19")) {
    $operations.Add(@{shortcut='+cells-merge';input=@{sheet_id='__SHEET__';range=$range;merge_type='all'}})
}
for ($index = $materialFirstIndex; $index -le $materialLastIndex; $index++) {
    $col = Get-ColName $index
    foreach ($range in @("${col}5:${col}6","${col}18:${col}19")) {
        $operations.Add(@{shortcut='+cells-merge';input=@{sheet_id='__SHEET__';range=$range;merge_type='all'}})
    }
}

$border = @{top=@{style='solid';color='#D9D9D9';weight='thin'};bottom=@{style='solid';color='#D9D9D9';weight='thin'};left=@{style='solid';color='#D9D9D9';weight='thin'};right=@{style='solid';color='#D9D9D9';weight='thin'}}
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range="A1:${lastCol}29";font_size=10;horizontal_alignment='center';vertical_alignment='middle';word_wrap='auto-wrap';border_styles=$border}})
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range="A1:${lastCol}3";background_color='#D9EAF7';font_weight='bold';font_size=12;horizontal_alignment='left'}})
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range="A4:${lastCol}4";background_color='#D9E2F3';font_weight='bold'}})
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range="A17:${lastCol}17";background_color='#D9E2F3';font_weight='bold'}})
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range='F9:F15';background_color='#F4CCCC';font_weight='bold'}})
$operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range='F23:F29';background_color='#F4CCCC';font_weight='bold'}})
foreach ($range in @("F10:${materialLastCol}11","F24:${materialLastCol}25")) {
    $operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range=$range;number_format='0.00%'}})
}
foreach ($range in @("F12:${materialLastCol}13","F26:${materialLastCol}27")) {
    $operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range=$range;number_format='#,##0.00'}})
}
foreach ($range in @("F14:${materialLastCol}15","F28:${materialLastCol}29")) {
    $operations.Add(@{shortcut='+cells-set-style';input=@{sheet_id='__SHEET__';range=$range;number_format='0.00'}})
}
foreach ($range in @("G8:${materialLastCol}8","G22:${materialLastCol}22")) {
    $operations.Add(@{shortcut='+dropdown-set';input=@{sheet_id='__SHEET__';range=$range;options=$DirectionOptions;multiple=$true;highlight=$true}})
}
$operations.Add(@{shortcut='+rows-resize';input=@{sheet_id='__SHEET__';range='1:29';type='pixel';size=28}})
$operations.Add(@{shortcut='+rows-resize';input=@{sheet_id='__SHEET__';range='5:6';type='pixel';size=58}})
$operations.Add(@{shortcut='+rows-resize';input=@{sheet_id='__SHEET__';range='18:19';type='pixel';size=58}})
$operations.Add(@{shortcut='+cols-resize';input=@{sheet_id='__SHEET__';range='A:F';type='pixel';size=110}})
$operations.Add(@{shortcut='+cols-resize';input=@{sheet_id='__SHEET__';range="G:${materialLastCol}";type='pixel';size=120}})
$operations.Add(@{shortcut='+cols-resize';input=@{sheet_id='__SHEET__';range=$lastCol;type='pixel';size=220}})
$operations.Add(@{shortcut='+dim-freeze';input=@{sheet_id='__SHEET__';dimension='column';count=6}})

if (-not $ConfirmWrite) {
    $dryOperations = @($operations | ForEach-Object {
        $copy = $_ | ConvertTo-Json -Depth 30 | ConvertFrom-Json
        $copy.input.sheet_id = 'dry-run-sheet'
        $copy
    })
    $createArgs = @('sheets','+workbook-create','--as','user','--title',$WorkbookTitle,'--dry-run')
    if ($FolderToken) { $createArgs += @('--folder-token',$FolderToken) }
    & lark-cli @createArgs | Out-Null
    $dryJson = ConvertTo-Json -InputObject ([object[]]$dryOperations) -Depth 30 -Compress
    Invoke-LarkWithInput @('sheets','+batch-update','--as','user','--spreadsheet-token','dry-run-token','--operations','-','--dry-run') $dryJson | Out-Null
    [ordered]@{ok=$true;mode='dry-run';workbook_title=$WorkbookTitle;plans=$plans.Count;material_capacity=$MaterialCapacity;columns=$noteIndex;note_column=$lastCol;operations=$operations.Count} | ConvertTo-Json
    exit 0
}

$createArgs = @('sheets','+workbook-create','--as','user','--title',$WorkbookTitle)
if ($FolderToken) { $createArgs += @('--folder-token',$FolderToken) }
$created = Invoke-LarkJson $createArgs
$spreadsheetToken = [string](Find-PropertyValue $created @('spreadsheet_token','token'))
$spreadsheetUrl = [string](Find-PropertyValue $created @('url','spreadsheet_url'))
if (-not $spreadsheetToken) { throw '工作簿已请求创建，但响应中没有 spreadsheet_token；请保留响应排查，不要重试创建' }

$info = Invoke-LarkJson @('sheets','+workbook-info','--as','user','--spreadsheet-token',$spreadsheetToken)
$defaultSheet = @($info.data.sheets)[0]
$defaultSheetId = [string]$defaultSheet.sheet_id
if (-not $defaultSheetId) { throw "工作簿已创建但找不到默认工作表：$spreadsheetToken" }
Invoke-LarkJson @('sheets','+sheet-rename','--as','user','--spreadsheet-token',$spreadsheetToken,'--sheet-id',$defaultSheetId,'--title',$TemplateTitle) | Out-Null

$actualOperations = @($operations | ForEach-Object {
    $copy = $_ | ConvertTo-Json -Depth 30 | ConvertFrom-Json
    $copy.input.sheet_id = $defaultSheetId
    $copy
})
$operationsJson = ConvertTo-Json -InputObject ([object[]]$actualOperations) -Depth 30 -Compress
$dryRun = Invoke-LarkWithInput @('sheets','+batch-update','--as','user','--spreadsheet-token',$spreadsheetToken,'--operations','-','--dry-run') $operationsJson
if ($dryRun -notmatch 'batch_update') { throw "模板 dry-run 未通过，工作簿已保留：$spreadsheetToken" }
Invoke-LarkWithInput @('sheets','+batch-update','--as','user','--spreadsheet-token',$spreadsheetToken,'--operations','-','--yes') $operationsJson | Out-Null

foreach ($plan in $plans) {
    Invoke-LarkJson @('sheets','+sheet-copy','--as','user','--spreadsheet-token',$spreadsheetToken,'--sheet-id',$defaultSheetId,'--title',([string]$plan.sheet_name)) | Out-Null
}
if ($IncludeMonthlySheet) {
    Invoke-LarkJson @('sheets','+sheet-create','--as','user','--spreadsheet-token',$spreadsheetToken,'--title','计划月数据','--row-count','200','--col-count','20') | Out-Null
}
Invoke-LarkJson @('sheets','+sheet-hide','--as','user','--spreadsheet-token',$spreadsheetToken,'--sheet-id',$defaultSheetId) | Out-Null

$finalInfo = Invoke-LarkJson @('sheets','+workbook-info','--as','user','--spreadsheet-token',$spreadsheetToken)
$sheetsByTitle = @{}
foreach ($sheet in @($finalInfo.data.sheets)) {
    $title = if ($sheet.sheet_name) { [string]$sheet.sheet_name } else { [string]$sheet.title }
    if ($sheetsByTitle.ContainsKey($title)) { throw "工作表标题重复：$title" }
    $sheetsByTitle[$title] = $sheet
}
$mappings = @()
foreach ($plan in $plans) {
    $title = [string]$plan.sheet_name
    if (-not $sheetsByTitle.ContainsKey($title)) { throw "创建后缺少工作表：$title" }
    $mappings += [ordered]@{campaign_id=[string]$plan.campaign_id;plan_name=[string]$plan.plan_name;sheet_id=[string]$sheetsByTitle[$title].sheet_id;sheet_name=$title}
}

[ordered]@{
    ok = $true
    mode = 'created_and_verified'
    spreadsheet_token = $spreadsheetToken
    spreadsheet_url = $spreadsheetUrl
    template_sheet_id = $defaultSheetId
    material_capacity = $MaterialCapacity
    note_column = $lastCol
    mappings = $mappings
} | ConvertTo-Json -Depth 8
