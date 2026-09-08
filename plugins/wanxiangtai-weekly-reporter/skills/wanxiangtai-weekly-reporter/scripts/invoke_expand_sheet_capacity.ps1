param(
    [Parameter(Mandatory = $true)][string]$SpreadsheetToken,
    [Parameter(Mandatory = $true)][string]$SheetId,
    [Parameter(Mandatory = $true)][string]$SheetTitle,
    [Parameter(Mandatory = $true)][string]$CurrentNoteCol,
    [Parameter(Mandatory = $true)][int]$RequiredMaterials
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Get-ColIndex([string]$Name) { $r=0; foreach($c in $Name.ToUpper().ToCharArray()){$r=$r*26+([int][char]$c-64)}; return $r }
function Get-ColName([int]$Index) { $r=''; while($Index -gt 0){$m=($Index-1)%26;$r=([char](65+$m))+$r;$Index=[math]::Floor(($Index-1)/26)}; return $r }
function Invoke-LarkJson([string[]]$Arguments) { $raw=(& lark-cli @Arguments 2>&1)-join "`n"; if($LASTEXITCODE -ne 0){throw "lark-cli failed: $raw"}; return $raw|ConvertFrom-Json }
function Invoke-LarkWithInput([string[]]$Arguments,[string]$InputText) { $raw=($InputText|& lark-cli @Arguments 2>&1)-join "`n"; if($LASTEXITCODE -ne 0){throw "lark-cli failed: $raw"}; return $raw }

$noteIndex=Get-ColIndex $CurrentNoteCol
# 素材从 G 列开始，备注列本身不计入容量，因此容量 = 备注列序号 - 7。
# 例如备注列 AB(28) 时，G:AA 共 21 个素材列。
$capacity=$noteIndex-7
$add=[math]::Max(0,$RequiredMaterials-$capacity)
if($add -eq 0){[pscustomobject]@{sheet=$SheetTitle;added=0;status='capacity_sufficient'}|ConvertTo-Json -Compress;exit 0}
if($noteIndex+$add -gt 200){throw '扩列后将超过飞书 200 列物理上限；不得截断，请创建素材续表'}
$sourceCol=Get-ColName($noteIndex-1)
$firstNew=Get-ColName($noteIndex)
$lastNew=Get-ColName($noteIndex+$add-1)
$newNote=Get-ColName($noteIndex+$add)

$workbook=Invoke-LarkJson @('sheets','+workbook-info','--as','user','--spreadsheet-token',$SpreadsheetToken)
$sheetInfo=@($workbook.data.sheets|Where-Object{[string]$_.sheet_id -eq [string]$SheetId})
if($sheetInfo.Count -ne 1){throw "无法唯一定位工作表：$SheetId"}
$rowCount=[int]$sheetInfo[0].row_count
if($rowCount -lt 29){throw "工作表行数不足：$rowCount"}
$pre=Invoke-LarkJson @('sheets','+sheet-info','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"A1:${CurrentNoteCol}${rowCount}",'--include','merges,col_widths')
$oldMergeCount=@($pre.data.merged_cells).Count
$oldNote=Invoke-LarkJson @('sheets','+csv-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"${CurrentNoteCol}1:${CurrentNoteCol}${rowCount}",'--rows-json')
$oldRows=@(); foreach($row in $oldNote.data.rows){$oldRows+=,"$($row.row_number)|$((@($row.values.psobject.Properties.Value)-join ''))"}

$ops=@(@{shortcut='+dim-insert';input=@{sheet_id=$SheetId;position=$CurrentNoteCol.ToUpper();count=$add;inherit_style='before'}})
for($i=$noteIndex;$i -lt $noteIndex+$add;$i++){$col=Get-ColName $i;$ops+=@{shortcut='+range-copy';input=@{sheet_id=$SheetId;source_range="${sourceCol}4:${sourceCol}29";target_range="${col}4";paste_type='formats'}}}
$ops+=@{shortcut='+cols-resize';input=@{sheet_id=$SheetId;range="${firstNew}:${lastNew}";type='pixel';size=105}}
$json=$ops|ConvertTo-Json -Depth 12 -Compress
$dry=Invoke-LarkWithInput @('sheets','+batch-update','--as','user','--spreadsheet-token',$SpreadsheetToken,'--operations','-','--dry-run') $json
if($dry -notmatch '"tool_name"\s*:\s*"batch_update"'){throw '扩列 dry-run 未通过'}
$raw=Invoke-LarkWithInput @('sheets','+batch-update','--as','user','--spreadsheet-token',$SpreadsheetToken,'--operations','-','--yes') $json
if(-not ($raw|ConvertFrom-Json).ok){throw $raw}

$layout=Invoke-LarkJson @('sheets','+sheet-info','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"A1:${newNote}${rowCount}",'--include','merges,col_widths')
$merges=@($layout.data.merged_cells|ForEach-Object{$_.range})
if(@($merges).Count -ne $oldMergeCount+2*$add){throw "合并数量不符：$(@($merges).Count)"}
foreach($required in @("A1:${newNote}1","A2:${newNote}2","A3:${newNote}3")){if($required -notin $merges){throw "缺少合并区 $required"}}
for($i=$noteIndex;$i -lt $noteIndex+$add;$i++){$col=Get-ColName $i;foreach($required in @("${col}5:${col}6","${col}18:${col}19")){if($required -notin $merges){throw "缺少合并区 $required"}}}

$newNoteData=Invoke-LarkJson @('sheets','+csv-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"${newNote}1:${newNote}${rowCount}",'--rows-json')
$newRows=@(); foreach($row in $newNoteData.data.rows){$newRows+=,"$($row.row_number)|$((@($row.values.psobject.Properties.Value)-join ''))"}
if(($oldRows -join "`n") -ne ($newRows -join "`n")){throw '备注列顺移后值不一致'}
$newCells=Invoke-LarkJson @('sheets','+csv-get','--as','user','--spreadsheet-token',$SpreadsheetToken,'--sheet-id',$SheetId,'--range',"${firstNew}4:${lastNew}${rowCount}",'--rows-json')
$nonempty=@($newCells.data.rows|ForEach-Object{$_.values.psobject.Properties.Value}|Where-Object{$_ -ne $null -and [string]$_ -ne ''})
if($nonempty.Count -gt 0){throw '新增素材列不应复制旧值'}
[pscustomobject]@{sheet=$SheetTitle;added=$add;new_note_col=$newNote;status='expanded_and_verified'}|ConvertTo-Json -Compress
