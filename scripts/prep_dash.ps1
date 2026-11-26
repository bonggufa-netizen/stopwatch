$sp = "$PSScriptRoot\..\work"
$out = "$sp\dashdata"; New-Item -ItemType Directory -Force $out | Out-Null
$master = Import-Csv "$sp\csv\매장마스터.csv" -Encoding UTF8
$code = @{}; foreach ($m in $master) { $code[$m.매장명] = $m.매장코드 }
function Get-Rnd($x, $n = 2) { [math]::Round([double]$x, $n) }
$sm = Import-Csv "$sp\sm.csv" -Encoding UTF8 | Sort-Object store, ym
$rows = foreach ($r in $sm) {
  [ordered]@{
    store = $r.store; store_code = $code[$r.store]; ym = $r.ym; year = [int]$r.yr; month = [int]$r.mo
    kind = $(if ($r.kind -eq '실적') { 'actual' } else { 'estimate' })
    sales = Get-Rnd $r.sales 1; orders = Get-Rnd $r.'총주문건수' 1; op_profit = Get-Rnd $r.'영업이익(만원)' 1
    food = Get-Rnd $r.'식자재비(만원)' 1; labor = Get-Rnd $r.'인건비(만원)' 1; rent = Get-Rnd $r.'임차료(만원)' 1
    platform_fee = Get-Rnd $r.'플랫폼수수료(만원)' 1; royalty = Get-Rnd $r.'로열티(만원)' 1; ad_fee = Get-Rnd $r.'광고분담금(만원)' 1; coupon = Get-Rnd $r.'쿠폰할인금액(만원)' 1
    delivery_share = Get-Rnd $r.'배달_주문비중' 4; hall_share = Get-Rnd $r.'홀_주문비중' 4; app_share = Get-Rnd $r.'자사앱주문비중' 4
    ext_orders = Get-Rnd ([double]$r.'배민_주문수' + [double]$r.'쿠팡이츠_주문수' + [double]$r.'요기요_주문수') 1
    qsc = Get-Rnd $r.'QSC점수' 1; hygiene = $r.hyg; claims = Get-Rnd $r.'클레임건수' 1; review = Get-Rnd $r.'리뷰평점' 2
    revisit = Get-Rnd $r.'재방문율' 4; waste = Get-Rnd $r.'식자재폐기율' 4; fulltime = Get-Rnd $r.'정규직원수' 2; parttime = Get-Rnd $r.'파트타임직원수' 2
  }
}
"store-month rows: $($rows.Count)"
[IO.File]::WriteAllText("$out\store_month.json", ($rows | ConvertTo-Json -Compress -Depth 3), [Text.UTF8Encoding]::new($false))

$alias = @{ '본사' = 'ALL'; '전 매장' = 'ALL'; '부평' = 'ST-004'; '충장로' = 'ST-008'; '둔산점' = 'ST-005' }
foreach ($m in $master) { $alias[$m.매장코드] = $m.매장코드; $alias[$m.'구매장코드(2022-08 이전)'] = $m.매장코드 }
$name = @{ 'ALL' = '전 매장·본사' }; foreach ($m in $master) { $name[$m.매장코드] = $m.매장명 }
$ev = Import-Csv "$sp\csv\사건로그.csv" -Encoding UTF8
$evs = foreach ($e in $ev) {
  $c = $alias[$e.'매장(기록 표기)']; if (-not $c) { throw "unmapped: $($e.'매장(기록 표기)')" }
  $d = [DateTime]::FromOADate([double]$e.일자)
  [ordered]@{ id = $e.사건ID; date = $d.ToString('yyyy-MM-dd'); ym = $d.ToString('yyyy-MM'); store_code = $c; store = $name[$c]; type = $e.유형; title = $e.제목; detail = $e.내용; status = $e.확인여부 }
}
"events: $($evs.Count)"
[IO.File]::WriteAllText("$out\events.json", ($evs | ConvertTo-Json -Compress -Depth 3), [Text.UTF8Encoding]::new($false))
Get-ChildItem $out | Select-Object Name, Length
