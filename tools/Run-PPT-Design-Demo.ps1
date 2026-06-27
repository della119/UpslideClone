<#
  Headless live-test of the PowerPoint design operations on REAL shapes, using the
  same tested UpslideClone.Core engines the ribbon buttons call. Builds a deck,
  applies Align / Distribute / Same-Size / TOC / Slide-Check, and asserts the result.
#>
$ErrorActionPreference = "Stop"
[Reflection.Assembly]::LoadFrom("D:\1_G\Upslide\src\UpslideClone.Core\bin\Debug\UpslideClone.Core.dll") | Out-Null
$Align = [UpslideClone.Core.Design.AlignEngine]
$AM = [UpslideClone.Core.Design.AlignMode]
$DA = [UpslideClone.Core.Design.DistributeAxis]
$SM = [UpslideClone.Core.Design.SizeMatch]
$pass = 0; $fail = 0
function Check($name, $cond) { if ($cond) { Write-Host "  PASS  $name" -ForegroundColor Green; $script:pass++ } else { Write-Host "  FAIL  $name" -ForegroundColor Red; $script:fail++ } }
function Boxes($shapes) { $l = New-Object "System.Collections.Generic.List[UpslideClone.Core.Design.LayoutBox]"; foreach ($s in $shapes) { $l.Add((New-Object UpslideClone.Core.Design.LayoutBox([single]$s.Left, [single]$s.Top, [single]$s.Width, [single]$s.Height))) }; return , $l }

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add(-1)

# cover + two titled slides (for TOC)
($pres.Slides.Add(1, 11)).Shapes.Title.TextFrame.TextRange.Text = "Cover"
($pres.Slides.Add(2, 11)).Shapes.Title.TextFrame.TextRange.Text = "Strategy"
($pres.Slides.Add(3, 11)).Shapes.Title.TextFrame.TextRange.Text = "Financials"
# a working slide with 4 mis-aligned rectangles
$ws = $pres.Slides.Add(4, 12)
$rects = @(
    $ws.Shapes.AddShape(1, 50, 50, 120, 40),
    $ws.Shapes.AddShape(1, 90, 150, 80, 60),
    $ws.Shapes.AddShape(1, 300, 30, 100, 50),
    $ws.Shapes.AddShape(1, 200, 260, 90, 70)
)
Write-Host "PowerPoint design live-test:" -ForegroundColor Cyan

# --- Smart Align (Left) ---
$r = $Align::Align((Boxes $rects), $AM::Left)
for ($i = 0; $i -lt 4; $i++) { $rects[$i].Left = $r[$i].Left; $rects[$i].Top = $r[$i].Top }
$minL = ($rects | ForEach-Object { $_.Left } | Measure-Object -Minimum).Minimum
Check "Smart Align Left -> all share min Left ($minL)" (@($rects | Where-Object { [math]::Abs($_.Left - $minL) -gt 0.5 }).Count -eq 0)

# --- Distribute (vertical) ---
$r = $Align::Distribute((Boxes $rects), $DA::Vertical)
for ($i = 0; $i -lt 4; $i++) { $rects[$i].Top = $r[$i].Top }
$sorted = $rects | Sort-Object { $_.Top }
$g1 = $sorted[1].Top - ($sorted[0].Top + $sorted[0].Height)
$g2 = $sorted[2].Top - ($sorted[1].Top + $sorted[1].Height)
Check "Distribute Down -> equal gaps ($([math]::Round($g1,1)) vs $([math]::Round($g2,1)))" ([math]::Abs($g1 - $g2) -lt 1.0)

# --- Same Size ---
$r = $Align::MatchSize((Boxes $rects), $SM::Both)
for ($i = 0; $i -lt 4; $i++) { $rects[$i].Width = $r[$i].Width; $rects[$i].Height = $r[$i].Height }
Check "Same Size -> all match first (W=$($rects[0].Width))" (@($rects | Where-Object { $_.Width -ne $rects[0].Width -or $_.Height -ne $rects[0].Height }).Count -eq 0)

# --- Table of Contents ---
$titles = New-Object "System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[int,string]]"
for ($i = 1; $i -le $pres.Slides.Count; $i++) { $t = ""; try { if ($pres.Slides.Item($i).Shapes.HasTitle -eq -1) { $t = $pres.Slides.Item($i).Shapes.Title.TextFrame.TextRange.Text } } catch {}; $titles.Add((New-Object "System.Collections.Generic.KeyValuePair[int,string]" ($i, $t))) }
$skip = New-Object "System.Collections.Generic.HashSet[int]"; [void]$skip.Add(1)
$toc = [UpslideClone.Core.Design.TableOfContents]::Build($titles, $skip)
Check "TOC built from titled slides (count=$($toc.Count), expect 2)" ($toc.Count -eq 2)

# --- Slide Check (push a shape off-slide, detect it) ---
$sw = $pres.PageSetup.SlideWidth; $sh = $pres.PageSetup.SlideHeight
$rects[0].Left = [single]($sw + 60)   # off the right edge
$box = New-Object UpslideClone.Core.Design.LayoutBox([single]$rects[0].Left, [single]$rects[0].Top, [single]$rects[0].Width, [single]$rects[0].Height)
$off = [UpslideClone.Core.Design.SlideCheckRules]::IsOffSlide($box, [single]$sw, [single]$sh, [single]1)
Check "Slide Check detects off-slide shape" ($off)

Write-Host ("Result: {0} passed, {1} failed." -f $pass, $fail) -ForegroundColor Cyan
Write-Host "PowerPoint left open with the test deck." -ForegroundColor Cyan
