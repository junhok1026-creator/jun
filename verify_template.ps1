$ErrorActionPreference = 'Stop'
$src = Join-Path (Get-Location) '청-지청-센터_자동취합_템플릿.xlsx'
$test = Join-Path (Get-Location) '청-지청-센터_자동취합_템플릿.test.xlsx'
Copy-Item -LiteralPath $src -Destination $test -Force

$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($test)

  $ws = $wb.Worksheets.Item('센터입력')
  $rows = @(
    @('A청','A지청','1센터','P001','마스크',10,''),
    @('A청','A지청','2센터','P001','마스크',20,''),
    @('A청','B지청','3센터','P001','마스크',5,''),
    @('A청','A지청','1센터','P002','장갑',7,''),
    @('A청','B지청','3센터','P003','소독제',3,'')
  )

  $r = 2
  foreach($data in $rows){
    for($c=1; $c -le 7; $c++){
      $val = $data[$c-1]
      if($c -eq 6){
        $ws.Cells.Item($r,$c).Value2 = [double]$val
      } else {
        $ws.Cells.Item($r,$c).Value2 = [string]$val
      }
    }
    $r++
  }

  $excel.CalculateFull()

  $org = $wb.Worksheets.Item('취합_청전체')
  $branch = $wb.Worksheets.Item('취합_지청별')
  $center = $wb.Worksheets.Item('취합_센터원본')

  $out = @()
  $out += "청전체 A2:C6"
  $out += ($org.Range('A2:C6').Value2 | Out-String)
  $out += "지청별 A2:D8"
  $out += ($branch.Range('A2:D8').Value2 | Out-String)
  $out += "센터원본 A2:E10"
  $out += ($center.Range('A2:E10').Value2 | Out-String)
  $out -join "`n"
}
finally {
  if($wb){$wb.Close($false); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)|Out-Null}
  if($excel){$excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)|Out-Null}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
