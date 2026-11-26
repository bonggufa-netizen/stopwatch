$sp = "$PSScriptRoot\..\work"
$d = Import-Csv "$sp\sm.csv" -Encoding UTF8 | Where-Object kind -eq '실적'
foreach ($r in $d) { foreach ($p in $r.PSObject.Properties) { if ($p.Name -notin 'store','ym','kind','hyg','inconsistent') { $p.Value = [double]$p.Value } } }
function S($rows, $f) { ($rows | Measure-Object -Property $f -Sum).Sum }
"### 배달 주문 중 외부앱 경유 비율 / 플랫폼수수료 ÷ 외부앱 주문 (원/건) / 자사앱 비중"
foreach ($y in 2020..2026) { $r = $d | Where-Object yr -eq $y
  $del = ($r | ForEach-Object { $_.'총주문건수' * $_.'배달_주문비중' } | Measure-Object -Sum).Sum
  $ext = (S $r '배민_주문수') + (S $r '쿠팡이츠_주문수') + (S $r '요기요_주문수')
  $app = ($r | ForEach-Object { $_.'총주문건수' * $_.'자사앱주문비중' } | Measure-Object -Sum).Sum
  "$y ext/deliv=" + ('{0:N1}' -f (100 * $ext / $del)) + "  ext/total=" + ('{0:N1}' -f (100 * $ext / (S $r '총주문건수'))) + "  fee/extorder=" + ('{0:N0}' -f (10000 * (S $r '플랫폼수수료(만원)') / $ext)) + "  app/total=" + ('{0:N1}' -f (100 * $app / (S $r '총주문건수'))) + "  baemin share of ext=" + ('{0:N1}' -f (100 * (S $r '배민_주문수') / $ext)) }
"### 광주 충장로점 2025-07~2026-08 월별"
$d | Where-Object { $_.store -eq '광주 충장로점' -and ($_.yr -eq 2026 -or ($_.yr -eq 2025 -and $_.mo -ge 7)) } | Sort-Object ym | ForEach-Object { "$($_.ym) sales=$($_.sales) opm=" + ('{0:N1}' -f (100 * $_.'영업이익(만원)' / $_.sales)) + " food%=" + ('{0:N1}' -f (100 * $_.'식자재비(만원)' / $_.sales)) + " order/food=" + ('{0:N2}' -f ($_.'식자재발주금액(만원)' / $_.'식자재비(만원)')) + " waste=$($_.'식자재폐기율') qsc=$($_.'QSC점수') ft=$($_.'정규직원수') pt=$($_.'파트타임직원수') labor%=" + ('{0:N1}' -f (100 * $_.'인건비(만원)' / $_.sales)) + " claim=$($_.'클레임건수') hyg=$($_.hyg)" }
"### 위생 B 등급 연도별 건수 (9개 매장 합)"
foreach ($y in 2020..2026) { "$y B=" + ($d | Where-Object { $_.yr -eq $y -and $_.hyg -eq 'B' }).Count + " / " + ($d | Where-Object yr -eq $y).Count }
"### 위생 B 월 vs A 월 평균 비교"
foreach ($h in 'A','B') { $r = $d | Where-Object hyg -eq $h; "$h n=$($r.Count) qsc=" + ('{0:N1}' -f ($r | Measure-Object QSC점수 -Average).Average) + " claims=" + ('{0:N1}' -f ($r | Measure-Object 클레임건수 -Average).Average) + " review=" + ('{0:N2}' -f ($r | Measure-Object 리뷰평점 -Average).Average) + " waste=" + ('{0:N2}' -f (100*($r | Measure-Object 식자재폐기율 -Average).Average)) }
"### 객단가 vs 메뉴 가격 — 2022 vs 2025 체인: 주문 1건당 피자 매출, 세트 매출"
foreach ($y in 2022, 2025, 2026) { $r = $d | Where-Object yr -eq $y; $o = S $r '총주문건수'; "$y pizza/order=" + ('{0:N0}' -f (10000 * (S $r 'c_피자') / $o)) + " set/order=" + ('{0:N0}' -f (10000 * (S $r 'c_세트메뉴') / $o)) + " total/order=" + ('{0:N0}' -f (10000 * (S $r 'sales') / $o)) }
"### 2026-03 신메뉴(트러플) 출시월 매장별 폐기율 2025-03 vs 2026-03"
$d | Where-Object { $_.mo -eq 3 -and $_.yr -ge 2025 } | Group-Object store | ForEach-Object { $a = $_.Group | Where-Object yr -eq 2025; $b = $_.Group | Where-Object yr -eq 2026; "$($_.Name) waste " + (100*$a.'식자재폐기율') + " -> " + (100*$b.'식자재폐기율') + "  new " + (100*$b.'신메뉴매출비중') }
