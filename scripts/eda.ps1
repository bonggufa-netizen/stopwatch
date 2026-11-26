$sp = "$PSScriptRoot\..\work"
$d = Import-Csv "$sp\sm.csv" -Encoding UTF8 | Where-Object kind -eq '실적'
foreach ($r in $d) { foreach ($p in $r.PSObject.Properties) { if ($p.Name -notin 'store','ym','kind','hyg','inconsistent') { $p.Value = [double]$p.Value } } }
function S($rows, $f) { ($rows | Measure-Object -Property $f -Sum).Sum }
function A($rows, $f) { ($rows | Measure-Object -Property $f -Average).Average }
function P($x) { '{0:N1}' -f $x }
# period label: years, with 2025 and 2026 compared on Jan-Aug
function Summ($rows, $label) {
  $s = S $rows 'sales'; $n = ($rows | Select-Object -ExpandProperty ym -Unique).Count
  [pscustomobject]@{
    period = $label; storeMonths = $rows.Count
    sales_m = [math]::Round($s / $n)                    # chain sales per month (만원)
    opm = P (100 * (S $rows '영업이익(만원)') / $s)
    food = P (100 * (S $rows '식자재비(만원)') / $s)
    labor = P (100 * (S $rows '인건비(만원)') / $s)
    rent = P (100 * (S $rows '임차료(만원)') / $s)
    platf = P (100 * (S $rows '플랫폼수수료(만원)') / $s)
    coupon = P (100 * (S $rows '쿠폰할인금액(만원)') / $s)
    orders_m = [math]::Round((S $rows '총주문건수') / $n)
    aov = '{0:N2}' -f ($s / (S $rows '총주문건수'))
    deliv = P (100 * (A $rows '배달_주문비중')); hall = P (100 * (A $rows '홀_주문비중')); app = P (100 * (A $rows '자사앱주문비중'))
    revisit = P (100 * (A $rows '재방문율')); qsc = P (A $rows 'QSC점수'); claims = P (A $rows '클레임건수'); review = '{0:N2}' -f (A $rows '리뷰평점')
    waste = P (100 * (A $rows '식자재폐기율'))
  }
}
"### 연도별 체인 요약 (월평균 매출, 비용은 매출 대비 %)"
$t = foreach ($y in 2020..2025) { Summ ($d | Where-Object yr -eq $y) "$y" }
$t += Summ ($d | Where-Object { $_.yr -eq 2025 -and $_.mo -le 8 }) '2025 1-8'
$t += Summ ($d | Where-Object { $_.yr -eq 2026 }) '2026 1-8'
$t | Format-Table -AutoSize | Out-String -Width 400

"### 매장별 2025 1-8 vs 2026 1-8"
$st = foreach ($g in ($d | Group-Object store)) {
  $a = $g.Group | Where-Object { $_.yr -eq 2025 -and $_.mo -le 8 }; $b = $g.Group | Where-Object { $_.yr -eq 2026 }
  $all = $g.Group | Where-Object { $_.yr -le 2025 }
  [pscustomobject]@{ store = $g.Name
    sales25 = [math]::Round((S $a 'sales') / 8); sales26 = [math]::Round((S $b 'sales') / 8)
    growth = P (100 * ((S $b 'sales') / (S $a 'sales') - 1))
    opm25 = P (100 * (S $a '영업이익(만원)') / (S $a 'sales')); opm26 = P (100 * (S $b '영업이익(만원)') / (S $b 'sales'))
    labor25 = P (100 * (S $a '인건비(만원)') / (S $a 'sales')); labor26 = P (100 * (S $b '인건비(만원)') / (S $b 'sales'))
    food25 = P (100 * (S $a '식자재비(만원)') / (S $a 'sales')); food26 = P (100 * (S $b '식자재비(만원)') / (S $b 'sales'))
    rent26 = P (100 * (S $b '임차료(만원)') / (S $b 'sales'))
    qsc26 = P (A $b 'QSC점수'); qscSD26 = P ([math]::Sqrt((($b | ForEach-Object { [math]::Pow($_.'QSC점수' - (A $b 'QSC점수'), 2) }) | Measure-Object -Sum).Sum / $b.Count))
    claim26 = P (A $b '클레임건수'); review26 = '{0:N2}' -f (A $b '리뷰평점')
    hygB_all = ($g.Group | Where-Object hyg -eq 'B').Count
    ft26 = P (A $b '정규직원수'); pt26 = P (A $b '파트타임직원수')
    deliv26 = P (100 * (A $b '배달_주문비중'))
    opm_2020_25 = P (100 * (S $all '영업이익(만원)') / (S $all 'sales'))
  }
}
$st | Sort-Object { [double]$_.opm26 } -Descending | Format-Table -AutoSize | Out-String -Width 400

"### 카테고리 믹스 (% of sales)"
$cats = 'c_피자','c_사이드메뉴','c_세트메뉴','c_음료','c_기타(굿즈/쿠폰)','c_배달수수료'
foreach ($y in 2020..2026) { $r = $d | Where-Object yr -eq $y; $s = S $r 'sales'; "$y  " + (($cats | ForEach-Object { "$($_.Substring(2))=" + (P (100 * (S $r $_) / $s)) }) -join '  ') }

"### 월별 체인 플랫폼수수료(매장평균, 만원)·배달비중 2024-08~2025-03"
foreach ($ym in '2024-08','2024-09','2024-10','2024-11','2024-12','2025-01','2025-02','2025-03') { $r = $d | Where-Object ym -eq $ym; "$ym  fee=" + [math]::Round((A $r '플랫폼수수료(만원)')) + "  fee%=" + (P (100 * (S $r '플랫폼수수료(만원)') / (S $r 'sales'))) + "  deliv=" + (P (100 * (A $r '배달_주문비중'))) + "  baemin=" + [math]::Round((A $r '배민_주문수')) }

"### 2026 월별 체인 (식자재율, 폐기율, 신메뉴, 인건비율, 영업이익률)"
foreach ($m in 1..8) { foreach ($y in 2025, 2026) { $r = $d | Where-Object { $_.yr -eq $y -and $_.mo -eq $m }; $s = S $r 'sales'; "$y-$m food=" + (P (100 * (S $r '식자재비(만원)') / $s)) + " waste=" + (P (100 * (A $r '식자재폐기율'))) + " new=" + (P (100 * (A $r '신메뉴매출비중'))) + " labor=" + (P (100 * (S $r '인건비(만원)') / $s)) + " opm=" + (P (100 * (S $r '영업이익(만원)') / $s)) + " sales=" + [math]::Round($s) } }

"### 손익 잔차 (매출 - 비용6 - 영업이익) vs 쿠폰, 연도별 % of sales"
foreach ($y in 2020..2026) { $r = $d | Where-Object yr -eq $y; $s = S $r 'sales'; $costs = 0; foreach ($k in '식자재비(만원)','인건비(만원)','임차료(만원)','플랫폼수수료(만원)','로열티(만원)','광고분담금(만원)') { $costs += S $r $k }; $res = $s - $costs - (S $r '영업이익(만원)'); "$y resid%=" + (P (100 * $res / $s)) + " coupon%=" + (P (100 * (S $r '쿠폰할인금액(만원)') / $s)) }

"### 상관 (매장·월, 실적 전체)"
function Corr($x, $y) { $mx = ($x | Measure-Object -Average).Average; $my = ($y | Measure-Object -Average).Average; $sxy = 0; $sxx = 0; $syy = 0; for ($i = 0; $i -lt $x.Count; $i++) { $a = $x[$i] - $mx; $b = $y[$i] - $my; $sxy += $a * $b; $sxx += $a * $a; $syy += $b * $b }; '{0:N2}' -f ($sxy / [math]::Sqrt($sxx * $syy)) }
foreach ($r in $d) { $r | Add-Member opm ([double]($r.'영업이익(만원)' / $r.sales * 100)) -Force; $r | Add-Member laborp ([double]($r.'인건비(만원)' / $r.sales * 100)) -Force; $r | Add-Member ptshare ([double]($r.'파트타임직원수' / ($r.'파트타임직원수' + $r.'정규직원수'))) -Force }
$pairs = @(('QSC점수','클레임건수'),('QSC점수','리뷰평점'),('클레임건수','리뷰평점'),('QSC점수','opm'),('ptshare','QSC점수'),('ptshare','laborp'),('ptshare','opm'),('리뷰평점','재방문율'),('자사앱주문비중','재방문율'),('배달_주문비중','opm'))
foreach ($p in $pairs) { "$($p[0]) ~ $($p[1]): " + (Corr @($d | ForEach-Object { [double]$_.($p[0]) }) @($d | ForEach-Object { [double]$_.($p[1]) })) }
