# Build store-month table from 매출데이터
$c = "$PSScriptRoot\..\work\csv"
$raw = Import-Csv "$c\매출데이터.csv" -Encoding UTF8
"raw rows: $($raw.Count)"
$g = $raw | Group-Object 매장명, '날짜(YYYY-MM)'
"store-months: $($g.Count)"
"groups not 6 rows: " + (($g | Where-Object Count -ne 6) | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '; '
$cols = '영업일수','총주문건수','홀_주문비중','배달_주문비중','포장_주문비중','배민_주문수','쿠팡이츠_주문수','요기요_주문수','점심피크_매출비중','저녁피크_매출비중','야간_매출비중','식자재비(만원)','인건비(만원)','임차료(만원)','플랫폼수수료(만원)','로열티(만원)','광고분담금(만원)','영업이익(만원)','신규고객수','재방문율','쿠폰할인금액(만원)','자사앱주문비중','리뷰평점','식자재발주금액(만원)','식자재폐기율','신메뉴매출비중','QSC점수','클레임건수','정규직원수','파트타임직원수'
$out = foreach ($x in $g) {
  $f = $x.Group[0]
  $o = [ordered]@{ store = $f.매장명; ym = $f.'날짜(YYYY-MM)'; yr = [int]$f.연도; mo = [int]$f.월; kind = $f.구분; hyg = $f.위생검사결과 }
  $tot = 0
  foreach ($r in $x.Group) { $v = [double]$r.'매출액(만원)'; $o['c_' + $r.매출카테고리] = $v; $tot += $v }
  $o['sales'] = $tot
  foreach ($k in $cols) { $o[$k] = [double]$f.$k }
  # consistency: same value across 6 rows?
  $o['inconsistent'] = (($cols | Where-Object { ($x.Group | Select-Object -ExpandProperty $_ -Unique).Count -gt 1 }) -join '|')
  [pscustomobject]$o
}
$out | Export-Csv "$c\..\sm.csv" -NoTypeInformation -Encoding UTF8
"inconsistent groups: " + ($out | Where-Object inconsistent).Count
$out | Group-Object kind | ForEach-Object { "$($_.Name): $($_.Count)" }
