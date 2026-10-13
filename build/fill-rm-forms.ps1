# Fills the "Retail Goal - Regional Manager" KPI form (one copy per RM) from build/out/rm-forms.json.
# Uses Excel itself so the form's formatting, merged cells and formulas are kept. Only writes cell values.
$outDir = "C:\Users\Lenovo\OneDrive\Desktop\Category Q1\26-27 iteamwise report H1\RM reports"
$data = Get-Content "C:\dev\PMS-H1-Review\build\out\rm-forms.json" -Raw | ConvertFrom-Json
$rowOf = @{ sales = 10; conv = 11; process = 12; salary = 13; team = 14 }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  foreach ($mgr in $data) {
    $path = Join-Path $outDir ("Retail Goal - Regional Manager - {0} - H1 FY 2026-27.xlsx" -f $mgr.name)
    $book = $excel.Workbooks.Open($path)
    $sheet = $book.Worksheets.Item("RM KPI")
    $sheet.Range("A1").Value2 = "Key Performance Indicators - 2026-27"
    $sheet.Range("C3").Value2 = $mgr.name
    $sheet.Range("C4").Value2 = "Retail Stores"
    $sheet.Range("C6").Value2 = "Praveen Kumar - Retail Business Head"
    $sheet.Range("C7").Value2 = "Mid Year Review: H1 FY 2026-27 (April to September 2026)"
    foreach ($kpi in $mgr.rows) {
      $line = $rowOf[$kpi.k]
      $remark = $sheet.Range("F$line")
      $achieved = $sheet.Range("H$line")
      switch ($kpi.k) {
        "sales"   { $remark.Value2 = "$($mgr.stores) stores. $($kpi.basis)"; $achieved.Value2 = [math]::Round([double]$mgr.H.ach, 1).ToString() + "%" }
        "conv"    { $remark.Value2 = "Conversion $($kpi.basis). UPT targets not available."; $achieved.Value2 = [math]::Round([double]$mgr.H.convAch, 1).ToString() + "%" }
        "team"    { $remark.Value2 = "$($kpi.basis)"; $achieved.Value2 = "Attrition $($kpi.ach)" }
        "process" { $remark.Value2 = "To be filled: list the initiatives taken in H1" }
        "salary"  { $remark.Value2 = "To be filled: salary to sales ratio for H1" }
      }
      if ($null -ne $kpi.r) {
        $sheet.Range("I$line").Value2 = [int]$kpi.r    # Self-Rating
        $sheet.Range("J$line").Value2 = [int]$kpi.r    # Manager's Rating (data-based; the weighted score formula uses this column)
      }
      $remark.WrapText = $true
    }
    $sheet.Columns.Item(6).ColumnWidth = 44
    [void]$sheet.Rows("10:14").AutoFit()
    $book.Save()
    Write-Output ("{0}: mid-year weighted score (K15) = {1}" -f $mgr.name, $sheet.Range("K15").Text)
    foreach ($line in 10..14) {
      Write-Output ("   {0} | ach {1} | self {2} | mgr {3} | weighted {4}" -f $sheet.Range("B$line").Text, $sheet.Range("H$line").Text, $sheet.Range("I$line").Text, $sheet.Range("J$line").Text, $sheet.Range("K$line").Text)
    }
    $book.Close($true)
  }
} finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
