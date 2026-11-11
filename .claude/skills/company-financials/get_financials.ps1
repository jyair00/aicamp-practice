<#
회사 요약재무 조회 (공공데이터포털 금융위원회_기업 재무정보 API)
사용 예:
  powershell -NoProfile -ExecutionPolicy Bypass -File get_financials.ps1 -Name "에스케이브로드밴드"
  powershell -NoProfile -ExecutionPolicy Bypass -File get_financials.ps1 -Crno 1101111466659 -FromYear 2020 -ToYear 2026
  -OutCsv 경로  → 결과를 CSV(UTF-8 BOM, 엑셀에서 바로 열림)로 저장
#>
param(
  [string]$Name,
  [string]$Crno,
  [int]$FromYear = ((Get-Date).Year - 5),
  [int]$ToYear = (Get-Date).Year,
  [string]$OutCsv
)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# 1) API 키: 환경변수 → 실습 폴더 .env
$key = $env:DATA_GO_KR_API_KEY
if (-not $key) {
  $envFile = Join-Path (Resolve-Path "$PSScriptRoot\..\..\..") '.env'
  if (Test-Path $envFile) {
    $line = Get-Content $envFile -Encoding UTF8 | Where-Object { $_ -match '^\s*DATA_GO_KR_API_KEY\s*=' } | Select-Object -First 1
    if ($line) { $key = ($line -split '=', 2)[1] }
  }
}
if (-not $key -or -not $key.Trim()) { Write-Output "[키 없음] DATA_GO_KR_API_KEY 를 환경변수나 .env 에 넣어주세요."; exit 2 }
$key = $key.Trim()

# 2) 회사명 → DART 고유번호 → 법인등록번호
if (-not $Crno) {
  if (-not $Name) { Write-Output "[입력 필요] -Name 또는 -Crno 를 주세요."; exit 2 }
  # DART 공시명은 회사마다 'SK텔레콤' 또는 '에스케이브로드밴드'처럼 표기가 달라 두 표기를 모두 시도
  $map = [ordered]@{ 'SK' = '에스케이'; 'LG' = '엘지'; 'KT' = '케이티'; 'GS' = '지에스'; 'CJ' = '씨제이'; 'HD' = '에이치디'; 'LS' = '엘에스' }
  $alt = $Name; foreach ($k in $map.Keys) { $alt = $alt -replace "(?i)^$k", $map[$k] }
  $alt2 = $Name; foreach ($k in $map.Keys) { $alt2 = $alt2 -replace "^$($map[$k])", $k }
  $tries = @($Name, $alt, $alt2) | Select-Object -Unique
  $cands = $null
  foreach ($q in $tries) {
    $enc = [uri]::EscapeDataString($q)
    $html = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 20 -Method Post 'https://dart.fss.or.kr/dsab007/detailSearch.ax' `
             -Body "textCrpNm=$enc&option=corp&currentPage=1&maxResults=50").Content
    $cands = [regex]::Matches($html, "openCorpInfoNew\('(\d{8})'[^>]*title=""([^""]+?) 기업개황") |
             ForEach-Object { [pscustomobject]@{ code = $_.Groups[1].Value; name = $_.Groups[2].Value.Trim() } } |
             Sort-Object code -Unique
    if ($cands) { $Name = $q; break }
  }
  if (-not $cands) { Write-Output "[회사 없음] DART에서 '$($tries -join "', '")' 을 찾지 못했어요. 공시 회사명으로 다시 시도하세요."; exit 3 }
  $pick = @($cands | Where-Object { $_.name -eq $Name })
  if ($pick.Count -ne 1) {
    if (@($cands).Count -eq 1) { $pick = @($cands) }
    else { Write-Output "[후보 여러 개] 정확한 이름이나 -Crno 로 다시 실행하세요:"; $cands | ForEach-Object { "  $($_.name) (DART $($_.code))" }; exit 4 }
  }
  $info = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 20 "https://dart.fss.or.kr/dsae001/selectPopup.ax?selectKey=$($pick[0].code)").Content -replace '<[^>]+>', ' ' -replace '\s+', ' '
  $m = [regex]::Match($info, '법인등록번호 (\d{6})-?(\d{7})')
  if (-not $m.Success) { Write-Output "[법인번호 없음] DART 기업개황에 법인등록번호가 없어요 ($($pick[0].name))."; exit 5 }
  $Crno = $m.Groups[1].Value + $m.Groups[2].Value
  $corpName = ([regex]::Match($info, '회사이름 (.+?) 정보 더보기')).Groups[1].Value
  Write-Output "[회사] $corpName / DART $($pick[0].code) / 법인등록번호 $($Crno.Substring(0,6))-$($Crno.Substring(6))"
}
$Crno = $Crno -replace '-', ''

# 3) 연도별 요약재무 조회 (연산명 끝에 _V2 필수)
$rows = @()
foreach ($y in $FromYear..$ToYear) {
  $u = "https://apis.data.go.kr/1160100/service/GetFinaStatInfoService_V2/getSummFinaStat_V2?serviceKey=$key&resultType=json&numOfRows=20&pageNo=1&crno=$Crno&bizYear=$y"
  try { $r = Invoke-RestMethod -Uri $u -TimeoutSec 20 }
  catch { Write-Output "[$y] 호출 오류 $($_.Exception.Response.StatusCode)"; continue }
  $rows += @($r.response.body.items.item) | Where-Object { $_ }
}
if (-not $rows) { Write-Output "[자료 없음] $FromYear~$ToYear 년 재무자료가 API에 없어요."; exit 6 }

$e = 100000000
function W($v) { if ($v -and $v -ne '') { [math]::Round([double]$v / $e) } else { $null } }
$table = $rows | Sort-Object fnclDcd, bizYear | ForEach-Object {
  [pscustomobject]@{
    구분 = ($_.fnclDcdNm -replace '요약재무제표', ''); 연도 = $_.bizYear
    매출 = W $_.enpSaleAmt; 영업이익 = W $_.enpBzopPft; 세전이익 = W $_.iclsPalClcAmt; 순이익 = W $_.enpCrtmNpf
    자산 = W $_.enpTastAmt; 부채 = W $_.enpTdbtAmt; 자본 = W $_.enpTcptAmt
    부채비율 = if ($_.fnclDebtRto) { [math]::Round([double]$_.fnclDebtRto, 1) } else { $null }
  }
}
Write-Output "[결과] $(@($rows).Count)건, 단위 억원(원 단위 반올림), 부채비율 %, 출처: 공공데이터포털 금융위원회_기업 재무정보"
$table | Format-Table -AutoSize | Out-String -Width 250
if ($OutCsv) { $table | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8; Write-Output "[저장] $OutCsv" }
