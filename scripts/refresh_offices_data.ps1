param(
  [string]$OutputJson = "data/offices.national.full.json",
  [string]$OutputJs = "data/offices.national.full.js",
  [string]$OverridesJson = "data/office_overrides.json",
  [int]$MinCenterCount = 100,
  [switch]$AllowLowCount
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
Set-Location $repoRoot

$base = "https://www.work.go.kr"
$seedUrl = "$base/geoje/main.do"
$sidoPattern = "(서울특별시|부산광역시|대구광역시|인천광역시|광주광역시|대전광역시|울산광역시|세종특별자치시|경기도|강원특별자치도|충청북도|충청남도|전북특별자치도|전라남도|경상북도|경상남도|제주특별자치도)"
$stdServices = @("실업급여", "국민취업지원", "취업알선", "직업훈련", "기업지원", "고용안정", "모성보호")
$serviceKeywordMap = [ordered]@{
  "실업급여" = @("실업급여", "구직급여", "실업인정", "수급자격")
  "국민취업지원" = @("국민취업지원", "구직촉진수당")
  "취업알선" = @("취업알선", "취업지원", "취업상담", "직업상담", "구인", "구직")
  "직업훈련" = @("직업훈련", "국민내일배움카드", "내일배움", "훈련")
  "기업지원" = @("기업지원", "고용장려금", "채용지원")
  "고용안정" = @("고용안정", "고용보험", "고용유지", "피보험", "이직확인서")
  "모성보호" = @("모성보호", "출산전후휴가", "육아휴직", "배우자출산")
  "외국인업무" = @("외국인", "외국인력", "고용허가제", "EPS")
}

function Clean([string]$s) {
  if ($null -eq $s) { return "" }
  return ($s -replace "\s+", " ").Trim()
}

function Get-ValueAfterLabel([string[]]$lines, [string]$label) {
  for ($i = $lines.Count - 1; $i -ge 0; $i--) {
    if ($lines[$i] -eq $label) {
      for ($j = $i + 1; $j -lt [Math]::Min($lines.Count, $i + 6); $j++) {
        $cand = Clean $lines[$j]
        if ($cand -and $cand -notmatch "^[\:\-]+$" -and $cand -ne $label) {
          return $cand
        }
      }
    }
  }
  for ($i = $lines.Count - 1; $i -ge 0; $i--) {
    if ($lines[$i] -like "$label*") {
      $rest = Clean ($lines[$i] -replace "^$([regex]::Escape($label))\s*", "")
      if ($rest) { return $rest }
    }
  }
  return ""
}

function Parse-HtmlLines([string]$html) {
  $text = [regex]::Replace($html, "<script[\s\S]*?</script>", " ", "IgnoreCase")
  $text = [regex]::Replace($text, "<style[\s\S]*?</style>", " ", "IgnoreCase")
  $text = [regex]::Replace($text, "<[^>]+>", "`n")
  $text = [System.Net.WebUtility]::HtmlDecode($text)
  return @($text -split "`n" | ForEach-Object { Clean $_ } | Where-Object { $_ })
}

function Infer-ServicesFromLines([string[]]$lines) {
  $found = @()
  foreach ($canon in $serviceKeywordMap.Keys) {
    foreach ($kw in $serviceKeywordMap[$canon]) {
      if ($lines -match [regex]::Escape($kw)) {
        $found += $canon
        break
      }
    }
  }
  return @($found | Sort-Object -Unique)
}

function Get-ServicesFromDeptPage([string]$centerId) {
  $deptUrl = "$base/$centerId/ctrIntro/deptStaffInfo/deptStaffInfoList.do?subNaviMenuCd=40300"
  try {
    $resp = Invoke-WebRequest -UseBasicParsing $deptUrl
    $lines = Parse-HtmlLines -html $resp.Content
    $services = Infer-ServicesFromLines -lines $lines
    return [pscustomobject]@{
      services = @($services)
      serviceVerified = [bool]($services.Count -gt 0)
      serviceSource = $deptUrl
    }
  } catch {
    return [pscustomobject]@{
      services = @()
      serviceVerified = $false
      serviceSource = $deptUrl
    }
  }
}

function Extract-Center([string]$href) {
  $centerId = ($href.Trim("/") -replace "/main\.do$", "")
  $url = "$base$href"
  $resp = Invoke-WebRequest -UseBasicParsing $url
  $html = $resp.Content

  $title = [regex]::Match($html, "<title>\s*(.*?)\s*</title>", "Singleline").Groups[1].Value
  $name = Clean ([System.Net.WebUtility]::HtmlDecode($title))
  if ($name -like "메인 -*") { $name = Clean $name.Substring(4) }

  $lines = Parse-HtmlLines -html $html
  $deptInfo = Get-ServicesFromDeptPage -centerId $centerId

  $jurText = Get-ValueAfterLabel -lines $lines -label "관할지역"

  $addrText = Get-ValueAfterLabel -lines $lines -label "주소"
  if ($addrText -notmatch "\(우\)") {
    $altAddr = ($lines | Where-Object { $_ -match "^\(우\)\s*\d+" } | Select-Object -Last 1)
    if ($altAddr) { $addrText = $altAddr }
  }
  $addrText = Clean ($addrText -replace "^\(우\)\s*\d+\s*", "")

  $telLine = ($lines | Where-Object { $_ -match "^TEL\s" } | Select-Object -Last 1)
  $tel = ""
  if ($telLine) { $tel = Clean (((($telLine -replace "^TEL\s*", "") -split "/")[0])) }

  $jurList = @()
  if ($jurText) {
    $jurList = $jurText -split "[,·/]" | ForEach-Object { Clean $_ } | Where-Object { $_ }
  }

  $sido = ""
  $sigungu = ""
  if ($addrText -match $sidoPattern) {
    $sido = $Matches[1]
    $rest = Clean ($addrText -replace "^$([regex]::Escape($sido))\s*", "")
    if ($rest -match "^([^\s]+(?:시|군|구))") { $sigungu = $Matches[1] }
  }
  if (-not $sigungu -and $jurList.Count -gt 0) {
    $sigungu = ($jurList[0] -split "\s+")[0]
  }

  $type = if ($name -like "*고용복지+센터*") { "고용복지+센터" } elseif ($name -like "*고용센터*") { "고용센터" } else { "고용센터" }
  $services = if ($deptInfo.serviceVerified) { @($deptInfo.services) } else { @($stdServices) }

  [pscustomobject]@{
    id = $centerId
    name = $name
    type = $type
    region = [pscustomobject]@{ sido = $sido; sigungu = $sigungu }
    jurisdiction = @($jurList)
    address = $addrText
    tel = $tel
    services = @($services)
    serviceVerified = $deptInfo.serviceVerified
    serviceSource = $deptInfo.serviceSource
    source = $url
  }
}

Write-Host "[1/4] Seed page fetch: $seedUrl"
$seed = Invoke-WebRequest -UseBasicParsing $seedUrl
$hrefsFromDom = @($seed.Links | ForEach-Object href)
$hrefsFromHtml = @([regex]::Matches($seed.Content, "/[a-zA-Z0-9_-]+/main\.do") | ForEach-Object { $_.Value })
$hrefs = @($hrefsFromDom + $hrefsFromHtml | Where-Object { $_ -match "^/[^/]+/main\.do$" } | Sort-Object -Unique)
Write-Host "Found center links: $($hrefs.Count)"
if ($hrefs.Count -lt 50) {
  throw "센터 링크 추출 건수가 비정상적으로 적습니다($($hrefs.Count)건). HTML 구조/접속환경을 확인하세요."
}

Write-Host "[2/4] Extracting center details..."
$centers = @()
foreach ($href in $hrefs) {
  try {
    $centers += Extract-Center -href $href
  } catch {
    Write-Warning "Failed: $href"
  }
}
if ((-not $AllowLowCount) -and $centers.Count -lt $MinCenterCount) {
  throw "센터 상세 수집 건수가 최소 기준($MinCenterCount)보다 적습니다. 현재: $($centers.Count)건. 기존 파일 덮어쓰기를 중단합니다."
}

Write-Host "[3/4] Applying overrides..."
$extras = @()
if (Test-Path $OverridesJson) {
  $loaded = Get-Content $OverridesJson -Raw | ConvertFrom-Json
  if ($loaded -is [System.Collections.IEnumerable]) {
    $extras = @($loaded)
  }
}

$merged = @{}
foreach ($r in $centers) { $merged[$r.id] = $r }
foreach ($e in $extras) {
  if (-not $e.id) { continue }
  if ($null -eq $e.PSObject.Properties["serviceVerified"]) {
    Add-Member -InputObject $e -NotePropertyName serviceVerified -NotePropertyValue $true
  }
  if ($null -eq $e.PSObject.Properties["serviceSource"]) {
    Add-Member -InputObject $e -NotePropertyName serviceSource -NotePropertyValue $e.source
  }
  $merged[$e.id] = $e
}

$all = $merged.Values | Sort-Object id

Write-Host "[4/4] Writing output files..."
$all | ConvertTo-Json -Depth 12 | Set-Content -Encoding UTF8 $OutputJson
$json = Get-Content $OutputJson -Raw
"window.NATIONAL_OFFICES = $json;" | Set-Content -Encoding UTF8 $OutputJs

$missingAddress = @($all | Where-Object { -not $_.address }).Count
$missingTel = @($all | Where-Object { -not $_.tel }).Count
$missingSido = @($all | Where-Object { -not $_.region.sido }).Count

Write-Host "Done"
Write-Host "Total records: $($all.Count)"
Write-Host "Missing address: $missingAddress"
Write-Host "Missing tel: $missingTel"
Write-Host "Missing sido: $missingSido"
