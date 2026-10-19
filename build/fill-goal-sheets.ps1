# Step 2 of the goal sheets: copies the HR "Retail Goal" form for each person in build/out/goal-sheets.json,
# fills the KPI rows and adds the formatted back-up sheets behind the KPI sheet. Run export-goal-sheets.js first.
param([string]$TemplateDir)
$forms = Get-Content "C:\dev\PMS-H1-Review\build\out\goal-sheets.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$headFill = 14998742   # light blue
$totalFill = 15921906  # light grey
try {
  foreach ($f in $forms) {
    Copy-Item (Join-Path $TemplateDir $f.template) $f.file -Force
    $book = $excel.Workbooks.Open($f.file)
    $kpi = $book.Worksheets.Item($f.kpiSheet)
    $kpi.Range("A1").Value2 = "Key Performance Indicators - 2026-27"
    $kpi.Range("C3").Value2 = $f.name
    $kpi.Range("C4").Value2 = "Retail Stores"
    $kpi.Range("C5").Value2 = $f.role
    $kpi.Range("C6").Value2 = $f.manager
    $kpi.Range("C7").Value2 = $f.period
    if ([string]::IsNullOrEmpty($kpi.Range("H9").Text)) { $kpi.Range("H9").Value2 = "Achievement %" }
    foreach ($r in $f.rows) {
      $n = $r.row
      $kpi.Range("F$n").Value2 = $r.remark
      $kpi.Range("F$n").WrapText = $true
      $kpi.Range("H$n").Value2 = $r.ach
      $kpi.Range("H$n").WrapText = $true
      if ($null -ne $r.r) {
        $kpi.Range("I$n").Value2 = [int]$r.r   # Self-Rating
        $kpi.Range("J$n").Value2 = [int]$r.r   # Manager's Rating (the weighted score formula reads this column)
      }
      if ($null -ne $r.m) { $kpi.Range("J$n").Value2 = [int]$r.m }   # changed on the dashboard by the manager
    }
    $kpi.Columns.Item(6).ColumnWidth = 46
    $kpi.Columns.Item(8).ColumnWidth = 16
    [void]$kpi.Rows("10:14").AutoFit()

    $src = $excel.Workbooks.Open($f.backup, 0, $true)
    $after = $kpi
    foreach ($s in $f.sheets) {
      $src.Worksheets.Item($s.name).Copy([Type]::Missing, $after)
      $after = $book.Worksheets.Item($after.Index + 1)
      $after.Cells.Font.Name = "Calibri"
      $after.Cells.Font.Size = 10
      $after.Range("A1").Font.Size = 13
      $after.Range("A1").Font.Bold = $true
      $after.Range("A2").Font.Italic = $true
      foreach ($t in $s.tables) {
        $box = $after.Range($after.Cells.Item($t.head, 1), $after.Cells.Item($t.last, $t.cols))
        $box.Borders.LineStyle = 1
        $box.Borders.Color = 12566463
        $head = $after.Range($after.Cells.Item($t.head, 1), $after.Cells.Item($t.head, $t.cols))
        $head.Font.Bold = $true
        $head.Interior.Color = $headFill
        $head.WrapText = $true
        $head.HorizontalAlignment = -4108
        $head.VerticalAlignment = -4108
        foreach ($tr in $t.totals) {
          $line = $after.Range($after.Cells.Item($tr, 1), $after.Cells.Item($tr, $t.cols))
          $line.Font.Bold = $true
          $line.Interior.Color = $totalFill
        }
        [void]$box.Columns.AutoFit()
        foreach ($c in 2..$t.cols) { if ($after.Columns.Item($c).ColumnWidth -lt 11) { $after.Columns.Item($c).ColumnWidth = 11 } }
        $after.Rows.Item($t.head).RowHeight = 30
        # the two lines under the table: the rating result and the scale
        $after.Cells.Item($t.last + 2, 1).Font.Bold = $true
        $after.Cells.Item($t.last + 2, 1).Font.Size = 11
        $after.Cells.Item($t.last + 3, 1).Font.Italic = $true
        $after.Activate()
        $excel.ActiveWindow.FreezePanes = $false
        $after.Cells.Item($t.head + 1, 2).Select()
        $excel.ActiveWindow.FreezePanes = $true
        $excel.ActiveWindow.DisplayGridlines = $false
      }
      if ($null -ne $s.mark) {   # shade the months a store achieved (above 90% of target)
        foreach ($c in $s.mark.cols) { foreach ($rw in $s.mark.first..$s.mark.last) {
          $cell = $after.Cells.Item($rw, $c)
          if ($cell.Value2 -is [double]) { if ($cell.Value2 -gt 0.9) { $cell.Interior.Color = 13561798; $cell.Font.Bold = $true } else { $cell.Interior.Color = 13551615 } }
        } }
      }
      $after.PageSetup.Orientation = 2
      $after.PageSetup.Zoom = $false
      $after.PageSetup.FitToPagesWide = 1
      $after.PageSetup.FitToPagesTall = $false
      $after.Cells.Item(1, 1).Select()
    }
    $src.Close($false)
    $kpi.Activate()
    $book.Save()
    Write-Output ("{0} - {1}: total {2} | {3}" -f $f.role, $f.name, $kpi.Range("K15").Text, (@($book.Worksheets | ForEach-Object { $_.Name }) -join ", "))
    foreach ($n in 10..14) { Write-Output ("     {0} | {1} | self {2} | mgr {3} | {4}" -f $kpi.Range("B$n").Text, $kpi.Range("H$n").Text, $kpi.Range("I$n").Text, $kpi.Range("J$n").Text, $kpi.Range("K$n").Text) }
    $book.Close($true)
  }
} finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
