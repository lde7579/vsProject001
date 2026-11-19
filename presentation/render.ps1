# 발표자료를 장별 PNG(1920x1080)로 내보낸다. PowerPoint(COM)가 설치된 PC에서 실행한다.
# 실행(저장소 루트에서): powershell -ExecutionPolicy Bypass -File presentation/render.ps1
param(
    [string]$Pptx = (Join-Path (Split-Path $PSScriptRoot -Parent) 'out\발표_적정재고일수.pptx'),
    [string]$OutDir = (Join-Path (Split-Path $PSScriptRoot -Parent) 'out\render')
)
$ErrorActionPreference = 'Stop'
$Pptx = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
# PowerPoint COM은 상대 경로를 읽지 못하므로 절대 경로로 바꾼다
$OutDir = (Resolve-Path $OutDir).Path
# 이전에 내보낸 장 그림만 지운다(다른 파일은 건드리지 않음)
Get-ChildItem -Path $OutDir -Filter 'slide-*.png' | ForEach-Object { Remove-Item -LiteralPath $_.FullName }

$app = New-Object -ComObject PowerPoint.Application
try {
    # 읽기 전용, 창 없이 연다
    $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
    try {
        $n = $pres.Slides.Count
        for ($i = 1; $i -le $n; $i++) {
            $file = Join-Path $OutDir ('slide-{0:D2}.png' -f $i)
            $pres.Slides.Item($i).Export($file, 'PNG', 1920, 1080)
        }
        Write-Output "$n 장을 $OutDir 에 내보냈습니다"
    } finally {
        $pres.Close()
    }
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
