param($d, $out)
$ns = @{m='http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
[xml]$ss = Get-Content "$d\xl\sharedStrings.xml" -Raw -Encoding UTF8
$strs = @()
foreach ($si in $ss.sst.si) { $strs += ($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '' }
[xml]$wb = Get-Content "$d\xl\workbook.xml" -Raw -Encoding UTF8
[xml]$rels = Get-Content "$d\xl\_rels\workbook.xml.rels" -Raw -Encoding UTF8
$map = @{}; foreach ($r in $rels.Relationships.Relationship) { $map[$r.Id] = $r.Target }
function ColIdx($ref) { $c = ($ref -replace '\d', ''); $n = 0; foreach ($ch in $c.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; $n - 1 }
foreach ($s in $wb.workbook.sheets.sheet) {
  $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $path = "$d\xl\" + ($map[$rid] -replace '^/xl/', '' -replace '/', '\')
  $reader = [System.Xml.XmlReader]::Create($path)
  $lines = New-Object System.Collections.Generic.List[string]
  $row = $null; $ref = $null; $t = $null; $val = $null; $inV = $false
  while ($reader.Read()) {
    if ($reader.NodeType -eq 'Element') {
      switch ($reader.LocalName) {
        'row' { $row = @{}; $max = -1 }
        'c' { $ref = $reader.GetAttribute('r'); $t = $reader.GetAttribute('t'); $val = ''; if ($reader.IsEmptyElement) { $i = ColIdx $ref; if ($i -gt $max) { $max = $i } } }
        'v' { $inV = $true }
        't' { if ($t -eq 'inlineStr') { $inV = $true } }
      }
    } elseif ($reader.NodeType -eq 'Text' -and $inV) { $val += $reader.Value }
    elseif ($reader.NodeType -eq 'EndElement') {
      switch ($reader.LocalName) {
        'v' { $inV = $false }
        't' { $inV = $false }
        'c' { if ($t -eq 's') { $val = $strs[[int]$val] }; $i = ColIdx $ref; $row[$i] = $val; if ($i -gt $max) { $max = $i } }
        'row' {
          $cells = for ($k = 0; $k -le $max; $k++) { $v = [string]$row[$k]; '"' + $v.Replace('"', '""') + '"' }
          $lines.Add(($cells -join ','))
        }
      }
    }
  }
  $reader.Close()
  [System.IO.File]::WriteAllLines("$out\$($s.name).csv", $lines, [System.Text.UTF8Encoding]::new($true))
  "$($s.name): $($lines.Count) rows"
}
