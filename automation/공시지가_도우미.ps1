# 공시지가 대량조회 도우미 (PowerShell판) — 파이썬 설치 없이 윈도우 내장 PowerShell만으로 동작
# 브라우저 HTML이 직접 못 부르는 브이월드 API를 127.0.0.1:43129 에서 대신 호출한다(외부 접속 불가).
# 사용:  powershell -ExecutionPolicy Bypass -File 공시지가_도우미.ps1 [-Mock]
# 키:    같은 폴더 vworld_key.txt — 1줄 인증키, 2줄 도메인(기본 localhost), 3줄 만료일(YYYY-MM-DD)
# 엔드포인트: /health · /landprice?pnu=&year= · /land?pnu= · /geocode?addr=   (파이썬판 공시지가_도우미.py 와 동일 규격)
param([switch]$Mock)
$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$VERSION = '2026.09.25.2-ps'
$PORT = 43129
$HERE = Split-Path -Parent $MyInvocation.MyCommand.Path
$KEY_FILE = Join-Path $HERE 'vworld_key.txt'
$VW_PRICE = 'https://api.vworld.kr/ned/data/getIndvdLandPriceAttr'
$VW_LAND  = 'https://api.vworld.kr/ned/data/ladfrlList'
$VW_GEO   = 'https://api.vworld.kr/req/address'
$MIN_GAP_MS = 150
$script:lastCall = [DateTime]::MinValue
$script:cache = @{}

function Load-Key {
    $key = ''; $domain = 'localhost'; $expiry = ''
    if (Test-Path $KEY_FILE) {
        $lines = Get-Content -LiteralPath $KEY_FILE -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
        if ($lines.Count -ge 1) { $key = $lines[0] }
        if ($lines.Count -ge 2 -and $lines[1]) { $domain = $lines[1] }
        if ($lines.Count -ge 3) { $expiry = $lines[2] }
    }
    return @{ key = $key; domain = $domain; expiry = $expiry }
}

# JSON 어디에 있든 '객체 배열'을 찾는다 (브이월드는 서비스마다 감싸는 키가 다르다)
function Find-List($obj) {
    if ($null -eq $obj) { return $null }
    if ($obj -is [System.Array]) {
        if ($obj.Count -gt 0 -and ($obj | Where-Object { -not ($_ -is [PSCustomObject]) }).Count -eq 0) { return $obj }
        foreach ($x in $obj) { $r = Find-List $x; if ($null -ne $r) { return $r } }
        return $null
    }
    if ($obj -is [PSCustomObject]) {
        foreach ($p in $obj.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [System.Array]) { $r = Find-List $v; if ($null -ne $r) { return $r } }
            elseif ($v -is [PSCustomObject]) { $r = Find-List $v; if ($null -ne $r) { return $r } }
        }
    }
    return $null
}
function Find-Key($obj, $name) {
    if ($obj -is [PSCustomObject]) {
        $p = $obj.PSObject.Properties[$name]; if ($p) { return $p.Value }
        foreach ($q in $obj.PSObject.Properties) { $r = Find-Key $q.Value $name; if ($null -ne $r) { return $r } }
    } elseif ($obj -is [System.Array]) { foreach ($x in $obj) { $r = Find-Key $x $name; if ($null -ne $r) { return $r } } }
    return $null
}
function Throttle { $gap = ([DateTime]::Now - $script:lastCall).TotalMilliseconds; if ($gap -lt $MIN_GAP_MS) { Start-Sleep -Milliseconds ([int]($MIN_GAP_MS - $gap)) }; $script:lastCall = [DateTime]::Now }
function Build-Url($base, $params) { $q = ($params.GetEnumerator() | ForEach-Object { [uri]::EscapeDataString([string]$_.Key) + '=' + [uri]::EscapeDataString([string]$_.Value) }) -join '&'; return "$base`?$q" }

function VWorld-Raw($url, $params) {
    $k = Load-Key
    if (-not $k.key) { return @{ data = $null; err = @{ ok = $false; error = 'NO_KEY'; message = 'vworld_key.txt 에 브이월드 인증키가 없습니다.' } } }
    $p = @{}; foreach ($e in $params.GetEnumerator()) { $p[$e.Key] = $e.Value }; $p['key'] = $k.key; $p['domain'] = $k.domain; $p['format'] = 'json'
    $full = Build-Url $url $p
    Throttle
    try { $raw = Invoke-WebRequest -Uri $full -UseBasicParsing -TimeoutSec 30 -Headers @{ 'User-Agent' = 'NawaLandPrice/1.0' }; $txt = $raw.Content }
    catch { return @{ data = $null; err = @{ ok = $false; error = 'NETWORK'; message = $_.Exception.Message } } }
    try { $data = $txt | ConvertFrom-Json } catch { return @{ data = $null; err = @{ ok = $false; error = 'BAD_JSON'; message = $txt.Substring(0, [Math]::Min(300, $txt.Length)) } } }
    return @{ data = $data; err = $null }
}
function VWorld-List($url, $params) {
    $ck = (Build-Url $url $params); if ($script:cache.ContainsKey($ck)) { return $script:cache[$ck] }
    $r = VWorld-Raw $url $params
    if ($r.err) { return $r.err }
    $code = Find-Key $r.data 'resultCode'; $msg = Find-Key $r.data 'resultMsg'; $items = Find-List $r.data
    if ($null -eq $items) { $items = @() }
    $src = ($url -split '/')[-1]
    $out = @{ ok = $true; items = @($items); total = (Find-Key $r.data 'totalCount'); resultCode = $code; resultMsg = $msg; source = $src }
    if ($code -and (@('OK','SUCCESS','NORMAL_SERVICE','00','0') -notcontains ([string]$code).ToUpper()) -and $items.Count -eq 0) {
        $out = @{ ok = $false; error = [string]$code; message = [string]$msg; source = $src }
    }
    if ($out.ok) { $script:cache[$ck] = $out }
    return $out
}
function Geocode($addr) {
    if ($script:cache.ContainsKey("geo|$addr")) { return $script:cache["geo|$addr"] }
    $common = @{ service = 'address'; version = '2.0'; crs = 'epsg:4326' }
    $p1 = $common.Clone(); $p1['request'] = 'getcoord'; $p1['address'] = $addr; $p1['refine'] = 'true'; $p1['simple'] = 'false'; $p1['type'] = 'road'
    $r = VWorld-Raw $VW_GEO $p1; if ($r.err) { return $r.err }
    $res = $r.data.response
    if ($res.status -ne 'OK') {
        $p2 = $p1.Clone(); $p2['type'] = 'parcel'
        $r2 = VWorld-Raw $VW_GEO $p2
        if (-not $r2.err -and $r2.data.response.status -eq 'OK') {
            $rf = $r2.data.response.refined
            return @{ ok = $true; source = 'geocoder'; jibun = [string]$rf.text; ldCode = [string]$rf.structure.level4LC; road = ''; point = $r2.data.response.result.point; note = '지번주소로 직접 인식' }
        }
        $m = ''; if ($res.error) { $m = [string]$res.error.text }; if (-not $m) { $m = '주소를 찾지 못했습니다.' }
        return @{ ok = $false; error = [string]$(if ($res.status) { $res.status } else { 'NOT_FOUND' }); message = $m; source = 'geocoder' }
    }
    $pt = $res.result.point; $roadText = [string]$res.refined.text; $detail = [string]$res.refined.structure.detail
    $p3 = $common.Clone(); $p3['request'] = 'getaddress'; $p3['point'] = "$($pt.x),$($pt.y)"; $p3['type'] = 'parcel'; $p3['zipcode'] = 'false'; $p3['simple'] = 'false'
    $r3 = VWorld-Raw $VW_GEO $p3; if ($r3.err) { return $r3.err }
    $res3 = $r3.data.response; $items = @($res3.result)
    if ($res3.status -ne 'OK' -or $items.Count -eq 0) { return @{ ok = $false; error = [string]$res3.status; message = '좌표에서 지번을 찾지 못했습니다.'; source = 'geocoder'; road = $roadText } }
    $it = $items[0]
    $out = @{ ok = $true; source = 'geocoder'; jibun = [string]$it.text; ldCode = [string]$it.structure.level4LC; bunji = [string]$it.structure.level5; road = $roadText; building = $detail; point = $pt; candidates = $items.Count }
    $script:cache["geo|$addr"] = $out; return $out
}
function Mock-Price($pnu, $year) {
    $seed = 0; foreach ($c in $pnu.ToCharArray()) { if ([char]::IsDigit($c)) { $seed += [int][string]$c } }
    $base = 500000 + (($seed * 7919) % 4500000); $y = [int]$year; $price = [int]($base * (1 + 0.03 * ($y - 2024)))
    return @{ ok = $true; mock = $true; source = 'mock'; items = @(@{ pnu = $pnu; stdrYear = "$y"; stdrMt = '01'; pblntfPclnd = "$price"; pblntfDe = "$y-04-30"; ldCodeNm = '모의 법정동'; lastUpdtDt = "$y-05-01" }) }
}
function Mock-Land($pnu) { $seed = 0; foreach ($c in $pnu.ToCharArray()) { if ([char]::IsDigit($c)) { $seed += [int][string]$c } }; return @{ ok = $true; mock = $true; source = 'mock'; items = @(@{ pnu = $pnu; lndcgrCodeNm = @('대','전','답','임야','공장용지')[$seed % 5]; lndpclAr = "$(100 + (($seed * 37) % 2900))"; posesnSeCodeNm = '개인' }) } }
function Mock-Geo($addr) { return @{ ok = $true; mock = $true; source = 'mock'; jibun = '서울특별시 강남구 역삼동 737'; ldCode = '1168010100'; bunji = '737'; road = $addr; building = '모의 건물'; candidates = 1 } }

# 쿼리 파라미터를 직접 UTF-8로 푼다 — Windows PowerShell 5.1(.NET Framework)의 HttpListener.QueryString 은 한글을 시스템 코드페이지로 풀어 깨진다
function Get-Q($req, $name) {
    $qs = $req.Url.Query; if (-not $qs) { return '' }
    foreach ($pair in $qs.TrimStart('?').Split('&')) {
        $i = $pair.IndexOf('='); if ($i -lt 0) { continue }
        if ($pair.Substring(0, $i) -eq $name) { return [uri]::UnescapeDataString($pair.Substring($i + 1).Replace('+', ' ')) }
    }
    return ''
}
function Send-Json($ctx, $obj, $status = 200) {
    $json = $obj | ConvertTo-Json -Depth 8 -Compress
    $b = [Text.Encoding]::UTF8.GetBytes($json)
    $r = $ctx.Response; $r.StatusCode = $status; $r.ContentType = 'application/json; charset=utf-8'
    $r.Headers.Add('Access-Control-Allow-Origin', '*'); $r.Headers.Add('Access-Control-Allow-Private-Network', 'true'); $r.Headers.Add('Cache-Control', 'no-store')
    $r.ContentLength64 = $b.Length; $r.OutputStream.Write($b, 0, $b.Length); $r.Close()
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://127.0.0.1:$PORT/")
try { $listener.Start() } catch { Write-Host "포트 $PORT 를 열지 못했습니다(이미 실행 중?): $($_.Exception.Message)"; exit 1 }
$k0 = Load-Key
Write-Host "공시지가 도우미(PowerShell) v$VERSION — http://127.0.0.1:$PORT  모의모드=$Mock 인증키=$(if ($k0.key) { '있음' } else { '없음' }) 도메인=$($k0.domain)"
while ($listener.IsListening) {
    try { $ctx = $listener.GetContext() } catch { break }
    try {
        $req = $ctx.Request; $path = $req.Url.AbsolutePath
        if ($req.HttpMethod -eq 'OPTIONS') { $r = $ctx.Response; $r.StatusCode = 204; $r.Headers.Add('Access-Control-Allow-Origin', '*'); $r.Headers.Add('Access-Control-Allow-Headers', 'Content-Type'); $r.Headers.Add('Access-Control-Allow-Methods', 'GET,OPTIONS'); $r.Headers.Add('Access-Control-Allow-Private-Network', 'true'); $r.Close(); continue }
        $k = Load-Key
        if ($path -eq '/health') { Send-Json $ctx @{ ok = $true; version = $VERSION; engine = 'powershell'; keyConfigured = [bool]$k.key; domain = $k.domain; mock = [bool]$Mock; keyFile = $KEY_FILE; keyExpiry = $k.expiry }; continue }
        $pnu = (Get-Q $req 'pnu').Trim()
        if ($path -eq '/landprice' -or $path -eq '/land') {
            if (-not ($pnu -match '^\d{19}$')) { Send-Json $ctx @{ ok = $false; error = 'BAD_PNU'; message = 'pnu는 19자리 숫자여야 합니다.' } 400; continue }
        }
        if ($path -eq '/landprice') {
            $year = (Get-Q $req 'year').Trim()
            if (-not ($year -match '^\d{4}$')) { Send-Json $ctx @{ ok = $false; error = 'BAD_YEAR'; message = 'year는 4자리 연도여야 합니다.' } 400; continue }
            $res = if ($Mock) { Mock-Price $pnu $year } else { VWorld-List $VW_PRICE @{ pnu = $pnu; stdrYear = $year; numOfRows = 50; pageNo = 1 } }
            Send-Json $ctx $res; continue
        }
        if ($path -eq '/land') { $res = if ($Mock) { Mock-Land $pnu } else { VWorld-List $VW_LAND @{ pnu = $pnu; numOfRows = 10; pageNo = 1 } }; Send-Json $ctx $res; continue }
        if ($path -eq '/geocode') {
            $addr = (Get-Q $req 'addr').Trim()
            if (-not $addr) { Send-Json $ctx @{ ok = $false; error = 'BAD_ADDR'; message = 'addr가 비어 있습니다.' } 400; continue }
            $res = if ($Mock) { Mock-Geo $addr } else { Geocode $addr }; Send-Json $ctx $res; continue
        }
        Send-Json $ctx @{ ok = $false; error = 'NOT_FOUND' } 404
    } catch { try { Send-Json $ctx @{ ok = $false; error = 'INTERNAL'; message = $_.Exception.Message } 500 } catch {} }
}
