[CmdletBinding()]
param([string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = "Stop"
$issues = [System.Collections.Generic.List[string]]::new()
function Test-RequiredFile([string]$RelativePath) {
    if (-not (Test-Path -LiteralPath (Join-Path $ProjectRoot $RelativePath) -PathType Leaf)) {
        $issues.Add("Eksik dosya: $RelativePath")
    }
}
Write-Host "OAA v2 - Kaynak Paket Kontrolü" -ForegroundColor Cyan
Write-Host "Proje: $ProjectRoot"
@("application/OAA.xlsm","src/modules/modConstants.bas","src/modules/modDatabase.bas","src/modules/modBootstrap.bas","src/modules/modTests.bas","src/classes/clsMember.cls","src/classes/clsMemberRepository.cls","src/classes/clsMemberService.cls") | ForEach-Object { Test-RequiredFile $_ }
$formRoot = Join-Path $ProjectRoot "src/forms"
if (Test-Path -LiteralPath $formRoot -PathType Container) {
    Get-ChildItem -LiteralPath $formRoot -Filter "*.frm" -File | ForEach-Object {
        $frx = [System.IO.Path]::ChangeExtension($_.FullName, ".frx")
        if (-not (Test-Path -LiteralPath $frx -PathType Leaf)) { $issues.Add("FRX eş dosyası eksik: $($_.Name)") }
    }
}
$testFile = Join-Path $ProjectRoot "src/modules/modTests.bas"
if (Test-Path -LiteralPath $testFile -PathType Leaf) {
    $testText = Get-Content -LiteralPath $testFile -Raw
    if ($testText -notmatch '(?im)\bTEST-144\b') { $issues.Add("modTests.bas içinde TEST-144 işareti bulunamadı.") }
    if ($testText -match '(?im)\bTEST-145\b') { Write-Host "TEST-145 kaynakta var; çalıştırma sonucu ayrıca doğrulanmalıdır." -ForegroundColor Yellow }
    else { Write-Host "TEST-145 henüz kaynakta yok (beklenen sonraki test)." -ForegroundColor Yellow }
}
if ($issues.Count -gt 0) {
    Write-Host "`nKAYNAK PAKETİ KONTROLÜ: BAŞARISIZ" -ForegroundColor Red
    $issues | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    exit 1
}
Write-Host "`nKAYNAK PAKETİ KONTROLÜ: BAŞARILI" -ForegroundColor Green
Write-Host "Not: Bu kontrol VBA Compile veya test çalıştırması değildir."
exit 0
