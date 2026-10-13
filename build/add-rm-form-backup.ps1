# Adds the back-up data sheets (from the RM performance workbook) behind the "RM KPI" sheet of each filled RM form.
$outDir = "C:\Users\Lenovo\OneDrive\Desktop\Category Q1\26-27 iteamwise report H1\RM reports"
$names = @("Manirathnam", "Sathish Kumar")
$backup = @("Summary", "Sales by month", "Conversion by month", "Store wise")
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  foreach ($name in $names) {
    $form = $excel.Workbooks.Open((Join-Path $outDir ("Retail Goal - Regional Manager - {0} - H1 FY 2026-27.xlsx" -f $name)))
    $perf = $excel.Workbooks.Open((Join-Path $outDir ("PMS H1 - RM & Cluster performance - {0}.xlsx" -f $name)), 0, $true)
    $have = @($form.Worksheets | ForEach-Object { $_.Name })
    $after = $form.Worksheets.Item("RM KPI")
    foreach ($s in $backup) {
      $target = "Backup - $s"
      if ($have -contains $target) { $after = $form.Worksheets.Item($target); continue }
      $perf.Worksheets.Item($s).Copy([Type]::Missing, $after)
      $after = $form.Worksheets.Item($after.Index + 1)
      $after.Name = $target
      [void]$after.Columns.AutoFit()
    }
    $form.Worksheets.Item("RM KPI").Activate()
    $form.Save()
    Write-Output ("{0}: {1}" -f $name, (@($form.Worksheets | ForEach-Object { "$($_.Name) [$($_.UsedRange.Rows.Count)r x $($_.UsedRange.Columns.Count)c]" }) -join " | "))
    $perf.Close($false)
    $form.Close($true)
  }
} finally {
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
