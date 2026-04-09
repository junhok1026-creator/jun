$ErrorActionPreference = 'Stop'

$outputPath = Join-Path (Get-Location) '청-지청-센터_자동취합_템플릿.xlsx'

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add(1)

    $wsInput = $workbook.Worksheets.Item(1)
    $wsInput.Name = '센터입력'

    $wsOrg = $workbook.Worksheets.Add()
    $wsOrg.Name = '취합_청전체'

    $wsBranch = $workbook.Worksheets.Add()
    $wsBranch.Name = '취합_지청별'

    $wsCenter = $workbook.Worksheets.Add()
    $wsCenter.Name = '취합_센터원본'

    $headers = @('청', '지청', '센터', '품목코드', '품목명', '필요수량', '비고(선택)', '집계키_청', '집계키_지청', '집계키_센터')
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $wsInput.Cells.Item(1, $i + 1).Value2 = $headers[$i]
    }

    $tblRange = $wsInput.Range('A1:J2')
    $listObj = $wsInput.ListObjects.Add(1, $tblRange, $null, 1)
    $listObj.Name = 'tbl_input'
    $listObj.TableStyle = 'TableStyleMedium2'

    $wsInput.Range('A1:J1').Font.Bold = $true

    $wsInput.Range('H2').Formula = '=[@품목코드]&"♦"&[@품목명]'
    $wsInput.Range('I2').Formula = '=[@지청]&"♦"&[@품목코드]&"♦"&[@품목명]'
    $wsInput.Range('J2').Formula = '=[@지청]&"♦"&[@센터]&"♦"&[@품목코드]&"♦"&[@품목명]'

    $requiredRules = @(
        @{ Col = 'A'; Formula = '=OR(COUNTA($A2:$F2)=0,LEN(TRIM(A2))>0)'; Title = '청 필수 입력'; Msg = '이 행을 입력할 때 청은 필수입니다.' },
        @{ Col = 'B'; Formula = '=OR(COUNTA($A2:$F2)=0,LEN(TRIM(B2))>0)'; Title = '지청 필수 입력'; Msg = '이 행을 입력할 때 지청은 필수입니다.' },
        @{ Col = 'C'; Formula = '=OR(COUNTA($A2:$F2)=0,LEN(TRIM(C2))>0)'; Title = '센터 필수 입력'; Msg = '이 행을 입력할 때 센터는 필수입니다.' },
        @{ Col = 'E'; Formula = '=OR(COUNTA($A2:$F2)=0,LEN(TRIM(E2))>0)'; Title = '품목명 필수 입력'; Msg = '이 행을 입력할 때 품목명은 필수입니다.' }
    )

    foreach ($rule in $requiredRules) {
        $rng = $wsInput.Range("$($rule.Col)2:$($rule.Col)1048576")
        $rng.Validation.Delete()
        $rng.Validation.Add(7, 1, 1, $rule.Formula)
        $rng.Validation.IgnoreBlank = $true
        $rng.Validation.InCellDropdown = $false
        $rng.Validation.ErrorTitle = $rule.Title
        $rng.Validation.ErrorMessage = $rule.Msg
        $rng.Validation.ShowError = $true
    }

    $qtyRange = $wsInput.Range('F2:F1048576')
    $qtyRange.Validation.Delete()
    $qtyRange.Validation.Add(7, 1, 1, '=OR(COUNTA($A2:$F2)=0,AND(ISNUMBER(F2),F2>=0))')
    $qtyRange.Validation.IgnoreBlank = $true
    $qtyRange.Validation.InCellDropdown = $false
    $qtyRange.Validation.ErrorTitle = '필요수량 숫자 입력'
    $qtyRange.Validation.ErrorMessage = '필요수량은 숫자만 입력할 수 있습니다. (0 이상)'
    $qtyRange.Validation.ShowError = $true
    $qtyRange.NumberFormat = '0'

    $wsInput.Columns.Item('H:J').Hidden = $true
    $wsInput.Range('A:G').EntireColumn.AutoFit() | Out-Null

    $maxRows = 5000

    $orgHeaders = @('품목코드', '품목명', '필요수량합계', '집계키(숨김)')
    for ($i = 0; $i -lt $orgHeaders.Count; $i++) {
        $wsOrg.Cells.Item(1, $i + 1).Value2 = $orgHeaders[$i]
    }
    $wsOrg.Range('A1:D1').Font.Bold = $true
    $wsOrg.Range("D2:D$maxRows").FormulaR1C1 = '=IFERROR(INDEX(tbl_input[집계키_청],ROW()-1),"")'
    $wsOrg.Range("A2:A$maxRows").FormulaR1C1 = '=IF(RC[3]="","",LEFT(RC[3],FIND("♦",RC[3])-1))'
    $wsOrg.Range("B2:B$maxRows").FormulaR1C1 = '=IF(RC[2]="","",MID(RC[2],FIND("♦",RC[2])+1,999))'
    $wsOrg.Range("C2:C$maxRows").FormulaR1C1 = '=IF(RC[1]="","",IF(COUNTIF(R2C4:RC[1],RC[1])>1,"",SUMIFS(tbl_input[필요수량],tbl_input[집계키_청],RC[1])))'
    $wsOrg.Columns.Item('D').Hidden = $true
    $wsOrg.Range("A1:C$maxRows").AutoFilter() | Out-Null
    $wsOrg.Range("A1:C$maxRows").AutoFilter(3, '<>') | Out-Null
    $wsOrg.Range('A:C').EntireColumn.AutoFit() | Out-Null

    $branchHeaders = @('지청', '품목코드', '품목명', '필요수량합계', '집계키(숨김)', 'p1', 'p2')
    for ($i = 0; $i -lt $branchHeaders.Count; $i++) {
        $wsBranch.Cells.Item(1, $i + 1).Value2 = $branchHeaders[$i]
    }
    $wsBranch.Range('A1:G1').Font.Bold = $true
    $wsBranch.Range("E2:E$maxRows").FormulaR1C1 = '=IFERROR(INDEX(tbl_input[집계키_지청],ROW()-1),"")'
    $wsBranch.Range("F2:F$maxRows").FormulaR1C1 = '=IF(RC[-1]="","",FIND("♦",RC[-1]))'
    $wsBranch.Range("G2:G$maxRows").FormulaR1C1 = '=IF(RC[-2]="","",FIND("♦",RC[-2],RC[-1]+1))'
    $wsBranch.Range("A2:A$maxRows").FormulaR1C1 = '=IF(RC[4]="","",LEFT(RC[4],RC[5]-1))'
    $wsBranch.Range("B2:B$maxRows").FormulaR1C1 = '=IF(RC[3]="","",MID(RC[3],RC[4]+1,RC[5]-RC[4]-1))'
    $wsBranch.Range("C2:C$maxRows").FormulaR1C1 = '=IF(RC[2]="","",MID(RC[2],RC[4]+1,999))'
    $wsBranch.Range("D2:D$maxRows").FormulaR1C1 = '=IF(RC[1]="","",IF(COUNTIF(R2C5:RC[1],RC[1])>1,"",SUMIFS(tbl_input[필요수량],tbl_input[집계키_지청],RC[1])))'
    $wsBranch.Columns.Item('E:G').Hidden = $true
    $wsBranch.Range("A1:D$maxRows").AutoFilter() | Out-Null
    $wsBranch.Range("A1:D$maxRows").AutoFilter(4, '<>') | Out-Null
    $wsBranch.Range('A:D').EntireColumn.AutoFit() | Out-Null

    $centerHeaders = @('지청', '센터', '품목코드', '품목명', '필요수량합계', '집계키(숨김)', 'p1', 'p2', 'p3')
    for ($i = 0; $i -lt $centerHeaders.Count; $i++) {
        $wsCenter.Cells.Item(1, $i + 1).Value2 = $centerHeaders[$i]
    }
    $wsCenter.Range('A1:I1').Font.Bold = $true
    $wsCenter.Range("F2:F$maxRows").FormulaR1C1 = '=IFERROR(INDEX(tbl_input[집계키_센터],ROW()-1),"")'
    $wsCenter.Range("G2:G$maxRows").FormulaR1C1 = '=IF(RC[-1]="","",FIND("♦",RC[-1]))'
    $wsCenter.Range("H2:H$maxRows").FormulaR1C1 = '=IF(RC[-2]="","",FIND("♦",RC[-2],RC[-1]+1))'
    $wsCenter.Range("I2:I$maxRows").FormulaR1C1 = '=IF(RC[-3]="","",FIND("♦",RC[-3],RC[-1]+1))'
    $wsCenter.Range("A2:A$maxRows").FormulaR1C1 = '=IF(RC[5]="","",LEFT(RC[5],RC[6]-1))'
    $wsCenter.Range("B2:B$maxRows").FormulaR1C1 = '=IF(RC[4]="","",MID(RC[4],RC[5]+1,RC[6]-RC[5]-1))'
    $wsCenter.Range("C2:C$maxRows").FormulaR1C1 = '=IF(RC[3]="","",MID(RC[3],RC[5]+1,RC[6]-RC[5]-1))'
    $wsCenter.Range("D2:D$maxRows").FormulaR1C1 = '=IF(RC[2]="","",MID(RC[2],RC[5]+1,999))'
    $wsCenter.Range("E2:E$maxRows").FormulaR1C1 = '=IF(RC[1]="","",IF(COUNTIF(R2C6:RC[1],RC[1])>1,"",SUMIFS(tbl_input[필요수량],tbl_input[집계키_센터],RC[1])))'
    $wsCenter.Columns.Item('F:I').Hidden = $true
    $wsCenter.Range("A1:E$maxRows").AutoFilter() | Out-Null
    $wsCenter.Range("A1:E$maxRows").AutoFilter(5, '<>') | Out-Null
    $wsCenter.Range('A:E').EntireColumn.AutoFit() | Out-Null

    $workbook.Worksheets.Item('센터입력').Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $workbook.Worksheets.Item('취합_청전체').Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $workbook.Worksheets.Item('취합_지청별').Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $workbook.Worksheets.Item('취합_센터원본').Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    $wsOrg.Protect()
    $wsBranch.Protect()
    $wsCenter.Protect()

    $xlOpenXMLWorkbook = 51
    $workbook.SaveAs($outputPath, $xlOpenXMLWorkbook)

    Write-Output "CREATED: $outputPath"
}
finally {
    if ($workbook) {
        $workbook.Close($true)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
