<#
  Headless demo runner — applies the Upslide engine's operations to the training
  files via Office automation, using the SAME tested UpslideClone.Core logic and
  theme the add-in uses. Leaves Excel + PowerPoint open with the results.
    W1: Smart Format, Waterfall, CAGR    W3: Export to PowerPoint (linked picture)
#>
$ErrorActionPreference = "Stop"
$core = "D:\1_G\Upslide\src\UpslideClone.Core\bin\Debug\UpslideClone.Core.dll"
$xlsx = "D:\1_G\Upslide\Training Guide Excel_test.xlsx"
$pptx = "D:\1_G\Upslide\Training Guide PPT_test.pptx"
[Reflection.Assembly]::LoadFrom($core) | Out-Null
$theme = [UpslideClone.Core.Branding.BrandTheme]::Default()

# Excel / Office constants
$xlContinuous = 1; $xlLeft = -4131; $xlRight = -4152
$xlEdgeLeft = 7; $xlEdgeTop = 8; $xlEdgeBottom = 9; $xlEdgeRight = 10
$xlColumnStacked = 52; $xlColumns = 2; $msoFalse = 0; $msoTrue = -1
$xlScreen = 1; $xlPicture = 2; $ppLayoutBlank = 12
$msoArrowTriangle = 2; $msoTextHorizontal = 1

function Ole($hex) { [UpslideClone.Core.Util.ColorUtil]::OleFromHex($hex) }
function DelShape($shapes, $name) { foreach ($s in @($shapes)) { if ($s.Name -eq $name) { $s.Delete() } } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsx)
Write-Host "Opened: $xlsx" -ForegroundColor Cyan

# ============ 1) SMART FORMAT — 'Format tables' (Income Statement B9:F<last>) ============
try {
    $ws = $wb.Worksheets.Item("Format tables")
    $top = 9; $left = 2; $cols = 5
    $r = $top + 1
    while (("" + $ws.Cells.Item($r, $left).Value2).Trim() -ne "") { $r++ }
    $last = $r - 1

    $rng = $ws.Range($ws.Cells.Item($top, $left), $ws.Cells.Item($last, $left + $cols - 1))
    $rng.Font.Name = $theme.Fonts.Latin; $rng.Font.Size = $theme.Fonts.SizeBody

    $hdr = $ws.Range($ws.Cells.Item($top, $left), $ws.Cells.Item($top, $left + $cols - 1))
    $hdr.Interior.Color = (Ole $theme.Colors.HeaderFill); $hdr.Font.Color = (Ole $theme.Colors.HeaderFont); $hdr.Font.Bold = $true
    $hdr.Borders.Item($xlEdgeBottom).LineStyle = $xlContinuous

    for ($c = 3; $c -le 5; $c++) { $ws.Range($ws.Cells.Item($top + 1, $c), $ws.Cells.Item($last, $c)).NumberFormat = $theme.NumberFormats.Number }
    $ws.Range($ws.Cells.Item($top + 1, 6), $ws.Cells.Item($last, 6)).NumberFormat = $theme.NumberFormats.Percent
    $ws.Range($ws.Cells.Item($top + 1, $left), $ws.Cells.Item($last, $left)).HorizontalAlignment = $xlLeft
    $ws.Range($ws.Cells.Item($top + 1, $left + 1), $ws.Cells.Item($last, $left + $cols - 1)).HorizontalAlignment = $xlRight

    $resultCount = 0
    for ($rr = $top + 1; $rr -le $last; $rr++) {
        $label = "" + $ws.Cells.Item($rr, $left).Value2
        if ([UpslideClone.Core.Formatting.SmartFormatRules]::IsResultRow($label)) {
            $row = $ws.Range($ws.Cells.Item($rr, $left), $ws.Cells.Item($rr, $left + $cols - 1))
            $row.Font.Bold = $true; $row.Interior.Color = (Ole $theme.Colors.ResultFill)
            $row.Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
            $resultCount++
        }
    }
    foreach ($edge in @($xlEdgeLeft, $xlEdgeTop, $xlEdgeBottom, $xlEdgeRight)) {
        $b = $rng.Borders.Item($edge); $b.LineStyle = $xlContinuous; $b.Color = (Ole $theme.Colors.Border)
    }
    Write-Host ("  [W1] Smart Format OK (green header, decimals; {0} result rows emphasised)." -f $resultCount) -ForegroundColor Green
} catch { Write-Host ("  Smart Format FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

# ============ 2) WATERFALL — 'Waterfall charts' (bridge B8:C13) ============
try {
    $ws2 = $wb.Worksheets.Item("Waterfall charts")
    DelShape $ws2.Shapes "UPS_WaterfallDemo"
    $points = New-Object "System.Collections.Generic.List[UpslideClone.Core.Charts.WaterfallPoint]"
    for ($rr = 8; $rr -le 13; $rr++) {
        $val = $ws2.Cells.Item($rr, 3).Value2
        if ($val -is [double]) {
            $p = New-Object UpslideClone.Core.Charts.WaterfallPoint
            $p.Label = [string]("" + $ws2.Cells.Item($rr, 2).Value2); $p.Value = [double]$val
            $points.Add($p)
        }
    }
    $rows = [UpslideClone.Core.Charts.WaterfallEngine]::Compute($points)

    $startRow = 30; $startCol = 2
    $c0 = $startCol; $c1 = $startCol + 1; $c2 = $startCol + 2; $c3 = $startCol + 3; $c4 = $startCol + 4
    $hh = @("Item", "Base", "Decrease", "Increase", "Total")
    for ($k = 0; $k -lt 5; $k++) { $ws2.Cells.Item($startRow, $startCol + $k).Value2 = [string]$hh[$k] }
    $i = 1
    foreach ($w in $rows) {
        $rr = $startRow + $i
        $ws2.Cells.Item($rr, $c0).Value2 = [string]$w.Label
        $ws2.Cells.Item($rr, $c1).Value2 = [double]$w.Base
        if ($null -eq $w.Decrease) { $ws2.Cells.Item($rr, $c2).Formula = "=NA()" } else { $ws2.Cells.Item($rr, $c2).Value2 = [double]$w.Decrease }
        if ($null -eq $w.Increase) { $ws2.Cells.Item($rr, $c3).Formula = "=NA()" } else { $ws2.Cells.Item($rr, $c3).Value2 = [double]$w.Increase }
        if ($null -eq $w.Total)    { $ws2.Cells.Item($rr, $c4).Formula = "=NA()" } else { $ws2.Cells.Item($rr, $c4).Value2 = [double]$w.Total }
        $i++
    }
    $helper = $ws2.Range($ws2.Cells.Item($startRow, $startCol), $ws2.Cells.Item($startRow + $rows.Count, $startCol + 4))
    $shp = $ws2.Shapes.AddChart2(-1, $xlColumnStacked); $shp.Name = "UPS_WaterfallDemo"
    $chart = $shp.Chart
    $chart.SetSourceData($helper, $xlColumns)
    $sc = $chart.SeriesCollection()
    for ($k = 1; $k -le $sc.Count; $k++) {
        $s = $sc.Item($k)
        switch ($s.Name) {
            "Base"     { $s.Format.Fill.Visible = $msoFalse }
            "Increase" { $s.Format.Fill.ForeColor.RGB = (Ole $theme.Colors.Increase); $s.HasDataLabels = $true }
            "Decrease" { $s.Format.Fill.ForeColor.RGB = (Ole $theme.Colors.Decrease); $s.HasDataLabels = $true; $s.DataLabels().NumberFormat = "(#,##0)" }
            "Total"    { $s.Format.Fill.ForeColor.RGB = (Ole $theme.Colors.Total);    $s.HasDataLabels = $true }
        }
    }
    $chart.HasTitle = $true; $chart.ChartTitle.Text = "Waterfall (Upslide)"
    $shp.Top = 60; $shp.Left = 380
    Write-Host "  [W1] Waterfall OK (dark-green anchors, red/green deltas)." -ForegroundColor Green
} catch { Write-Host ("  Waterfall FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

# ============ 3) CAGR ARROW — 'CAGR Arrow' (first chart's first series) ============
try {
    $wsC = $wb.Worksheets.Item("CAGR Arrow")
    if ($wsC.ChartObjects().Count -lt 1) { throw "no embedded chart on the sheet" }
    $cchart = $wsC.ChartObjects(1).Chart
    $svals = $cchart.SeriesCollection(1).Values
    $list = New-Object "System.Collections.Generic.List[double]"
    foreach ($v in $svals) { if ($v -is [double]) { $list.Add([double]$v) } }
    $cagr = [UpslideClone.Core.Charts.CagrEngine]::Compute($list)

    DelShape $cchart.Shapes "UPS_CAGR_Arrow"
    DelShape $cchart.Shapes "UPS_CAGR_Label"
    $pa = $cchart.PlotArea
    $x0 = $pa.InsideLeft; $y0 = $pa.InsideTop; $pw = $pa.InsideWidth; $ph = $pa.InsideHeight
    $arrow = $cchart.Shapes.AddLine([single]($x0 + $pw * 0.10), [single]($y0 + $ph * 0.85), [single]($x0 + $pw * 0.90), [single]($y0 + $ph * 0.20))
    $arrow.Name = "UPS_CAGR_Arrow"
    $arrow.Line.EndArrowheadStyle = $msoArrowTriangle; $arrow.Line.Weight = 2.0
    $arrow.Line.ForeColor.RGB = (Ole $theme.Colors.Total)
    $tb = $cchart.Shapes.AddTextbox($msoTextHorizontal, [single]($x0 + $pw * 0.38), [single]($y0 + $ph * 0.06), 130, 24)
    $tb.Name = "UPS_CAGR_Label"
    $tb.TextFrame.Characters().Text = $cagr.Label
    $tb.TextFrame.Characters().Font.Bold = $true
    $tb.Line.Visible = $msoFalse; $tb.Fill.Visible = $msoFalse
    Write-Host ("  [W1] CAGR OK ({0} over {1} periods)." -f $cagr.Label, $cagr.Periods) -ForegroundColor Green
} catch { Write-Host ("  CAGR skipped/failed: " + $_.Exception.Message) -ForegroundColor Yellow }

# ============ 2b) STACKED WATERFALL — 'Stacked Waterfall' (B8:E14, FR/UK/DE) ============
try {
    $wsS = $wb.Worksheets.Item("Stacked Waterfall")
    DelShape $wsS.Shapes "UPS_StackedDemo"
    $catArr = New-Object "string[]" 3
    for ($c = 0; $c -lt 3; $c++) { $catArr[$c] = [string]("" + $wsS.Cells.Item(8, 3 + $c).Value2) }
    $pts = New-Object "System.Collections.Generic.List[UpslideClone.Core.Charts.StackedWaterfallPoint]"
    for ($rr = 9; $rr -le 14; $rr++) {
        $vals = New-Object "double[]" 3
        for ($c = 0; $c -lt 3; $c++) { $vv = $wsS.Cells.Item($rr, 3 + $c).Value2; if ($vv -is [double]) { $vals[$c] = [double]$vv } else { $vals[$c] = 0 } }
        $p = New-Object UpslideClone.Core.Charts.StackedWaterfallPoint
        $p.Label = [string]("" + $wsS.Cells.Item($rr, 2).Value2); $p.Values = $vals
        $pts.Add($p)
    }
    $res = [UpslideClone.Core.Charts.StackedWaterfallEngine]::Compute($catArr, $pts)
    $sr = 30; $sc = 2; $bcols = 2 + 3
    $hh2 = @("Item", "Base") + $catArr
    for ($k = 0; $k -lt $bcols; $k++) { $wsS.Cells.Item($sr, $sc + $k).Value2 = [string]$hh2[$k] }
    $i = 1
    foreach ($w in $res.Rows) {
        $rr = $sr + $i
        $wsS.Cells.Item($rr, $sc).Value2 = [string]$w.Label
        $wsS.Cells.Item($rr, $sc + 1).Value2 = [double]$w.Base
        for ($c = 0; $c -lt 3; $c++) { if ($null -eq $w.Segments[$c]) { $wsS.Cells.Item($rr, $sc + 2 + $c).Formula = "=NA()" } else { $wsS.Cells.Item($rr, $sc + 2 + $c).Value2 = [double]$w.Segments[$c] } }
        $i++
    }
    $helperS = $wsS.Range($wsS.Cells.Item($sr, $sc), $wsS.Cells.Item($sr + $res.Rows.Count, $sc + $bcols - 1))
    $shpS = $wsS.Shapes.AddChart2(-1, $xlColumnStacked); $shpS.Name = "UPS_StackedDemo"
    $chS = $shpS.Chart; $chS.SetSourceData($helperS, $xlColumns)
    $palette = @($theme.Colors.Increase, $theme.Colors.Total, $theme.Colors.Decrease)
    $scS = $chS.SeriesCollection(); $ci = 0
    for ($k = 1; $k -le $scS.Count; $k++) {
        $s = $scS.Item($k)
        if ($s.Name -eq "Base") { $s.Format.Fill.Visible = $msoFalse } else { $s.Format.Fill.ForeColor.RGB = (Ole $palette[$ci % 3]); $ci++ }
    }
    $chS.HasTitle = $true; $chS.ChartTitle.Text = "Stacked Waterfall (Upslide)"; $shpS.Top = 60; $shpS.Left = 430
    Write-Host ("  [W1] Stacked Waterfall OK ({0} steps x 3 categories)." -f $res.Rows.Count) -ForegroundColor Green
} catch { Write-Host ("  Stacked Waterfall FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

# ============ 5) AUTOCOLOR — 'Advanced Export' WACC model (C11:H23) ============
try {
    $wsA = $wb.Worksheets.Item("Advanced Export")
    $blue = 0; $black = 0; $green = 0
    for ($rr = 11; $rr -le 23; $rr++) {
        for ($c = 3; $c -le 8; $c++) {
            $cell = $wsA.Cells.Item($rr, $c)
            $hasF = [bool]$cell.HasFormula
            $isEmpty = ($null -eq $cell.Value2) -and (-not $hasF)
            $formula = $null; if ($hasF) { $formula = [string]$cell.Formula }
            $cls = [UpslideClone.Core.Modelling.AutocolorClassifier]::Classify($hasF, $formula, $isEmpty)
            if ($cls -ne [UpslideClone.Core.Modelling.CellColorClass]::Empty) {
                $cell.Font.Color = (Ole ([UpslideClone.Core.Modelling.AutocolorClassifier]::DefaultHex($cls)))
                if ($cls -eq [UpslideClone.Core.Modelling.CellColorClass]::Input) { $blue++ } elseif ($cls -eq [UpslideClone.Core.Modelling.CellColorClass]::Formula) { $black++ } else { $green++ }
            }
        }
    }
    Write-Host ("  [W2] Autocolor OK ({0} inputs blue, {1} formulas black, {2} links green)." -f $blue, $black, $green) -ForegroundColor Green
} catch { Write-Host ("  Autocolor FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

# ============ 6) IFERROR — wrap the market-value formulas (Advanced Export C18:H20) ============
try {
    $wrapped = 0
    for ($rr = 18; $rr -le 20; $rr++) {
        for ($c = 3; $c -le 8; $c++) {
            $cell = $wsA.Cells.Item($rr, $c)
            if ([bool]$cell.HasFormula) {
                $f = [string]$cell.Formula
                $u = [UpslideClone.Core.Modelling.FormulaTransform]::WrapIfError($f, '""')
                if ($u -ne $f) { $cell.Formula = $u; $wrapped++ }
            }
        }
    }
    Write-Host ("  [W2] IFERROR OK ({0} formulas wrapped)." -f $wrapped) -ForegroundColor Green
} catch { Write-Host ("  IFERROR FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

$wb.Save()
Write-Host "Saved workbook." -ForegroundColor Cyan

# ============ 4) EXPORT TO POWERPOINT — formatted Income Statement as a linked picture ============
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $pres = $ppt.Presentations.Open($pptx)
    $slide = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)

    $src = $ws.Range($ws.Cells.Item(8, 2), $ws.Cells.Item($last, 6))   # Income Statement incl. title row 8
    $src.CopyPicture($xlScreen, $xlPicture)
    $pasted = $slide.Shapes.Paste()
    $shape = $pasted.Item(1)
    $shape.Left = 40; $shape.Top = 70

    $meta = New-Object UpslideClone.Core.Linking.LinkMetadata
    $meta.LinkId = [UpslideClone.Core.Linking.LinkMetadata]::NewId()
    $meta.SourceWorkbook = $xlsx; $meta.SourceSheet = "Format tables"; $meta.SourceRange = "`$B`$8:`$F`$$last"
    $meta.ExportType = [UpslideClone.Core.Linking.ExportType]::Picture
    $meta.LastRefresh = [DateTime]::Now
    foreach ($kv in $meta.ToTags().GetEnumerator()) { $shape.Tags.Add($kv.Key, [string]$kv.Value) }

    $pres.Save()
    Write-Host ("  [W3] Export to PowerPoint OK (linked picture on slide {0}, tagged UPS_LinkId)." -f $slide.SlideIndex) -ForegroundColor Green
} catch { Write-Host ("  Export to PowerPoint FAILED at line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message) -ForegroundColor Red }

Write-Host "Done. Excel + PowerPoint left open with results." -ForegroundColor Cyan
