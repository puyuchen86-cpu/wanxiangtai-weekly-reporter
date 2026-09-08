param(
    [Parameter(Mandatory = $true)][string]$PlanJson,
    [Parameter(Mandatory = $true)][string]$HistoryRoot,
    [string]$ManualDirectionsJson = '{}',
    [Parameter(Mandatory = $true)][string]$BestImageId,
    [string]$BestVideoId = ''
)

$ErrorActionPreference = 'Stop'
$plan = Get-Content -LiteralPath $PlanJson -Raw | ConvertFrom-Json
$history = @{}
$allowedDirections = @{
    '用户信任' = $true
    '圈人群' = $true
    '正向' = $true
    '反向' = $true
    '功能' = $true
    '情感价值' = $true
    '使用感受' = $true
    '使用结果' = $true
}
Get-ChildItem -LiteralPath $HistoryRoot -Recurse -File -Filter '*-enrichment.json' |
    Where-Object { $_.FullName -ne [IO.Path]::ChangeExtension($PlanJson, $null) + '-enrichment.json' } |
    ForEach-Object {
        $entry = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
        foreach ($property in $entry.directions.psobject.Properties) {
            $values = @($property.Value)
            $valid = $values.Count -gt 0
            foreach ($value in $values) {
                if (-not $allowedDirections.ContainsKey([string]$value)) { $valid = $false; break }
            }
            if ($valid -and -not $history.ContainsKey($property.Name)) { $history[$property.Name] = $values }
        }
    }
$manualObject = $ManualDirectionsJson | ConvertFrom-Json
$manual = @{}
foreach ($property in $manualObject.psobject.Properties) { $manual[$property.Name] = @($property.Value) }

$directions = [ordered]@{}
foreach ($item in $plan.materials) {
    $id = [string]$item.id
    if ($manual.ContainsKey($id)) { $directions[$id] = $manual[$id] }
    elseif ($history.ContainsKey($id)) { $directions[$id] = $history[$id] }
    else { throw "缺少方向：$id / $($item.name)" }
}

$output = [ordered]@{directions=$directions;best_image_id=$BestImageId;best_video_id=if($BestVideoId){$BestVideoId}else{$null}}
$output | ConvertTo-Json -Depth 12
