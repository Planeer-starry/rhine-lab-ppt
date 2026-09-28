# -*- coding: utf-8 -*-
<#
  莱茵生命 RHINE LAB 风格渲染器
  用法：
    $code = [IO.File]::ReadAllText('<本文件>', [Text.Encoding]::UTF8); Invoke-Expression $code
    # 或带参数：
    # 在调用前设置 $env:RHINE_CONTENT / $env:RHINE_PPTX / $env:RHINE_PNG

  内容来源：-Content 指定的 JSON；未指定时用内置母版示例。
  设计语言完全固化在本脚本中，改内容不改风格。
#>

$ErrorActionPreference = 'Stop'

# ---------------- 参数 ----------------
$ContentPath = $env:RHINE_CONTENT
if (-not $ContentPath) { $ContentPath = 'D:\DeepSeek Harness\大学事项\ppt模板\莱茵生命-母版内容.json' }
$PptxOut = $env:RHINE_PPTX
if (-not $PptxOut) { $PptxOut = 'D:\DeepSeek Harness\大学事项\ppt模板\莱茵生命-RHINE-LAB-风格母版.pptx' }
$PngDir = $env:RHINE_PNG
if (-not $PngDir) { $PngDir = 'D:\DeepSeek Harness\大学事项\_rhine\master' }

New-Item -ItemType Directory -Force -Path (Split-Path $PptxOut -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path $PngDir | Out-Null
Get-ChildItem "$PngDir\*.png" -ErrorAction SilentlyContinue | Remove-Item -Force

# ---------------- 设计 token（BGR 十进制） ----------------
$INK   = 0x151515   # #151515 墨黑
$IVORY = 0xE3EAED   # #EDEAE3 暖象牙
$DARK  = 0x0C0B0B   # #0B0B0C 深场
$GOLD  = 0x5FA1C0   # #C0A15F 香槟金
$WHITE = 0xFFFFFF
$GHOST = 0x1A1A1A   # 深场幽灵字
$GREY  = 0x666A6B   # #6B6A66
$GREY2 = 0x858A8C   # #8C8A85
$RULE  = 0xBCD2D3   # #D3D2BC 浅发丝线
$CARD  = 0x171715

$F_ZH_B = '思源黑体 CN Bold'
$F_ZH_H = '思源黑体 CN Heavy'
$F_ZH_N = '思源黑体 CN Normal'
$F_ZH_R = '思源黑体 CN Regular'
$F_EN   = 'Bahnschrift'
$F_LB   = 'Arial Narrow'

$W = 960.0; $H = 540.0
$Cx0 = 48.0; $Cx1 = 912.0; $CW = 864.0     # 版心

# ---------------- 基础绘制 ----------------
# PowerPoint COM 无 Font.Spacing，用发丝空格 U+200A 模拟宽字距
function Trk {
    param([string]$Text,[double]$LS=0)
    if ($LS -le 0 -or $Text -eq '') { return $Text }
    $pad = [string]::new([char]0x200A, [Math]::Max(1, [int][Math]::Round($LS / 3.0)))
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        [void]$sb.Append($ch)
        if ($ch -eq "`n") { continue }
        [void]$sb.Append($pad)
    }
    return $sb.ToString()
}

function Txt {
    param($Slide,$Name,$X,$Y,$W,$H,$Text,$Font,$Size,$Color,
          [double]$LS=0,[double]$Sp=1.0,[int]$Align=1,[int]$Anchor=1)
    $sh = $Slide.Shapes.AddTextbox(1, $X, $Y, $W, $H)
    $sh.Name = $Name
    $tf = $sh.TextFrame
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.VerticalAnchor = $Anchor
    $tr = $tf.TextRange
    $tr.Text = (Trk $Text $LS)
    $tr.Font.Name = $Font
    $tr.Font.NameFarEast = $Font
    $tr.Font.Size = [single]$Size
    $tr.Font.Color.RGB = [int]$Color
    $tr.Font.Bold = 0
    $tr.ParagraphFormat.Alignment = $Align
    $tr.ParagraphFormat.SpaceWithin = $Sp
    $tr.ParagraphFormat.SpaceBefore = 0
    $tr.ParagraphFormat.SpaceAfter = 0
    return $sh
}

function Rect {
    param($Slide,$Name,$X,$Y,$W,$H,$Color,[double]$Opacity=1.0,[int]$Shape=1)
    $sh = $Slide.Shapes.AddShape($Shape, $X, $Y, $W, $H)
    $sh.Name = $Name
    if ($Opacity -lt 1.0) {
        $sh.Fill.Transparency = (1.0 - $Opacity)
    } else {
        $sh.Fill.Transparency = 0
    }
    $sh.Fill.ForeColor.RGB = $Color
    $sh.Fill.Visible = -1
    $sh.Line.Visible = 0
    $sh.Shadow.Visible = 0
    return $sh
}

function Rule {
    param($Slide,$Name,$X,$Y,$W,$Color,[double]$Thick=1.0,[double]$Opacity=1.0)
    return (Rect $Slide $Name $X $Y $W $Thick $Color $Opacity)
}

function Add-Slide {
    param([int]$Bg=$IVORY)
    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)   # ppLayoutBlank
    Rect $s 'bg' 0 0 $W $H $Bg | Out-Null
    return $s
}

# 按目标比例预裁出临时图（像素级，绝不拉伸）
function Crop-ToRatio {
    param([string]$Path,[double]$TargetAR,[double]$Bias=0.5)
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create((New-Object System.Uri($Path)), 'None', 'Default')
    $frame = $dec.Frames[0]
    $iw = [double]$frame.PixelWidth; $ih = [double]$frame.PixelHeight
    $ar = $iw / $ih
    if ([math]::Abs($ar - $TargetAR) -lt 0.005) { return $Path }   # 已匹配，直接用原图
    if ($ar -gt $TargetAR) {
        $cw = [int][math]::Round($ih * $TargetAR); $ch = [int]$ih
        $cx = [int][math]::Round(($iw - $cw) * 0.5); $cy = 0
    } else {
        $cw = [int]$iw; $ch = [int][math]::Round($iw / $TargetAR)
        $cx = 0; $cy = [int][math]::Round(($ih - $ch) * $Bias)
    }
    if ($cw -lt 1) { $cw = 1 }; if ($ch -lt 1) { $ch = 1 }
    if ($cx -lt 0) { $cx = 0 }; if ($cy -lt 0) { $cy = 0 }
    if ($cx + $cw -gt $iw) { $cx = [int]$iw - $cw }
    if ($cy + $ch -gt $ih) { $cy = [int]$ih - $ch }
    $rect = New-Object System.Windows.Int32Rect $cx, $cy, $cw, $ch
    $cropped = New-Object System.Windows.Media.Imaging.CroppedBitmap($frame, $rect)
    $dir = Join-Path ([System.IO.Path]::GetTempPath()) 'rhine-precrop'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $out = Join-Path $dir ([System.IO.Path]::GetFileNameWithoutExtension($Path) + '-' + [int]($TargetAR*1000) + '.png')
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($cropped))
    $fs = [System.IO.File]::Create($out)
    $enc.Save($fs); $fs.Close()
    return $out
}

# 图片：预裁到目标比例 → 等比缩放到精确框位。因为比例已一致，赋值宽高不会变形。
function Pic {
    param($Slide,$Name,$Path,[double]$X,[double]$Y,[double]$W,[double]$H,[double]$CropBias=0.5)
    if (-not (Test-Path -LiteralPath $Path)) { throw "图片不存在: $Path" }
    $prepped = Crop-ToRatio -Path $Path -TargetAR ($W / $H) -Bias $CropBias
    $sh = $Slide.Shapes.AddPicture($prepped, 0, -1, [single]$X, [single]$Y, [single]$W, [single]$H)
    $sh.Name = $Name
    $sh.Left = $X; $sh.Top = $Y
    $sh.Width = $W; $sh.Height = $H
    $sh.Line.Visible = 0
    $sh.Shadow.Visible = 0
    return $sh
}

function Wordmark {
    param($Slide,[int]$Color=$INK)
    Txt $Slide 'mark1' 48 34 320 34 'RHINE LAB' $F_EN 24 $Color 1 1.05 1 1 | Out-Null
    Txt $Slide 'mark2' 48 64 320 20 'SYNTHESIZE  INFORMATION' $F_LB 10 $Color 2.2 1.2 1 1 | Out-Null
    Txt $Slide 'mark3' 48 86 320 22 'ANALYSIS   OS' $F_LB 14 $Color 3 1.2 1 1 | Out-Null
}

function Footer {
    param($Slide,[int]$Color=$INK,[string]$Page='')
    Txt $Slide 'footer' 640 492 210 20 'POWERED BY RHINE LAB' $F_LB 9 $Color 1.6 1 3 1 | Out-Null
    Rect $Slide 'footerbar' 866 496 46 11 $Color | Out-Null
    if ($Page -ne '') { Txt $Slide 'pageno' 640 512 210 16 $Page $F_LB 9 $Color 1.6 1 3 1 | Out-Null }
}

function Head {
    param($Slide,$Section,$Title,$Kicker)
    Rect $Slide 'secbar' 48 142 96 21 $INK | Out-Null
    Txt $Slide 'sectext' 48 142 96 21 $Section $F_LB 9 $WHITE 1.4 1 2 3 | Out-Null
    Txt $Slide 'title' 48 176 720 54 $Title $F_ZH_H 34 $INK 2 1 1 3 | Out-Null
    if ($Kicker -and $Kicker -ne '') {
        Txt $Slide 'kicker' 700 176 212 54 $Kicker $F_LB 10 $GREY 1.6 1 3 3 | Out-Null
    }
    Rule $Slide 'headrule' 48 248 864 $INK 1 | Out-Null
}

# ---------------- 版式 ----------------
function L-Cover {
    param($Slide,$C)
    Rect $Slide 'ringO' 612 122 398 398 $WHITE 0.55 9 | Out-Null
    Rect $Slide 'ringI' 660 170 302 302 $WHITE 0.34 9 | Out-Null
    Wordmark $Slide $INK
    Rect $Slide 'ctick' 356 244 26 26 $GOLD | Out-Null
    Txt $Slide 'ctitle' 400 226 470 62 ([string]$C.title) $F_ZH_B 50 $INK 4 1 1 3 | Out-Null
    Txt $Slide 'csub' 400 308 470 26 ([string]$C.subtitle) $F_LB 12 $GREY 2.5 1 1 1 | Out-Null
    Footer $Slide $INK
}

function L-Agenda {
    param($Slide,$C)
    Wordmark $Slide $INK
    Txt $Slide 'akick' 48 150 400 20 ([string]$C.kicker) $F_LB 10 $GREY 3 1 1 1 | Out-Null
    Txt $Slide 'atitle' 48 176 500 56 ([string]$C.title) $F_ZH_H 38 $INK 3 1 1 3 | Out-Null
    Rule $Slide 'arule' 48 246 864 $INK 1 | Out-Null
    $n = $C.items.Count
    $gap = 62.0
    if ($n -ge 4) { $gap = 52.0 }
    if ($n -ge 6) { $gap = 42.0 }
    $y = 278; $i = 0
    foreach ($it in $C.items) {
        $i++
        Txt $Slide ('no'+$i) 48 $y 60 30 ([string]$it.no) $F_EN 22 $GOLD 1 1 1 3 | Out-Null
        Txt $Slide ('zh'+$i) 128 $y 420 30 ([string]$it.zh) $F_ZH_B 20 $INK 0 1 1 3 | Out-Null
        Txt $Slide ('en'+$i) 560 ($y+3) 352 26 ([string]$it.en) $F_LB 10 $GREY 2.2 1 3 1 | Out-Null
        $y += $gap
        Rule $Slide ('r'+$i) 48 ($y-12) 864 $RULE 1 | Out-Null
    }
    Footer $Slide $INK
}

function L-Process {
    param($Slide,$C)
    Wordmark $Slide $INK
    Txt $Slide 'fkick' 48 150 400 20 ([string]$C.kicker) $F_LB 10 $GREY 3 1 1 1 | Out-Null
    Txt $Slide 'ftitle' 48 176 700 56 ([string]$C.title) $F_ZH_H 34 $INK 2 1 1 3 | Out-Null
    Rule $Slide 'frule' 48 248 864 $INK 1 | Out-Null
    Rule $Slide 'fline' 196 298 495 $INK 1 | Out-Null
    $xs = @(48, 300, 552, 804); $ws = @(240, 240, 240, 108)
    $i = 0
    foreach ($st in $C.stages) {
        $x = [double]$xs[$i]; $wd = [double]$ws[$i]
        Txt $Slide ('no'+$i) $x 272 100 22 ([string]$st.no) $F_EN 18 $GOLD 1 1 1 3 | Out-Null
        Txt $Slide ('ti'+$i) $x 314 $wd 30 ([string]$st.title) $F_ZH_B 19 $INK 0 1 1 3 | Out-Null
        Txt $Slide ('de'+$i) $x 348 $wd 66 ([string]$st.desc) $F_ZH_N 12 $GREY 0 1.45 1 1 | Out-Null
        $cx = $x + 120 - 7
        Rect $Slide ('nd'+$i) $cx 292 13 13 $INK | Out-Null
        $i++
    }
    if ($C.readout) { Txt $Slide 'fread' 48 440 864 34 ([string]$C.readout) $F_ZH_N 16 $INK 0 1.4 1 1 | Out-Null }
    if ($C.note)    { Txt $Slide 'fnote' 48 486 620 20 ([string]$C.note) $F_LB 10 $GREY 1.4 1 1 1 | Out-Null }
    Footer $Slide $INK
}

function L-Dark {
    param($Slide,$C)
    $gridYs = @(170, 300, 462)
    $i = 0
    foreach ($gy in $gridYs) { $i++; Rule $Slide ('gh'+$i) 0 $gy 960 $WHITE 1 0.93 | Out-Null }
    for ($i=1; $i -le 4; $i++) { Rect $Slide ('gv'+$i) ($i*180) 0 1 540 $WHITE 0.07 | Out-Null }
    Txt $Slide 'dghost' 600 120 340 80 'RHINE LAB' $F_EN 52 $GHOST 2 1 3 3 | Out-Null
    Wordmark $Slide $WHITE
    Txt $Slide 'dtitle' 48 200 800 76 ([string]$C.title) $F_ZH_R 42 $WHITE 3 1 1 3 | Out-Null
    Txt $Slide 'dsub' 48 284 800 26 ([string]$C.subtitle) $F_LB 11 $GREY2 2.6 1 1 1 | Out-Null
    Rule $Slide 'dline' 202 376 180 $WHITE 1 | Out-Null
    Rect $Slide 'danchor' 382 372 10 10 $WHITE | Out-Null
    if ($C.label) {
        Rect $Slide 'dlabel' 412 363 94 21 $WHITE | Out-Null
        Txt $Slide 'dlabeltxt' 412 363 94 21 ([string]$C.label) $F_LB 9 $INK 1.4 1 2 3 | Out-Null
    }
    if ($C.card) {
        Rect $Slide 'dcard' 431 399 224 116 $CARD | Out-Null
        Rule $Slide 'dcardrule' 451 419 184 $WHITE 1 0.5 | Out-Null
        Txt $Slide 'dcardtxt' 451 429 184 70 ($C.card -join "`n") $F_LB 11 $WHITE 0.6 1.7 1 1 | Out-Null
    }
    if ($C.caption) {
        Txt $Slide 'dcap' 48 326 800 24 ([string]$C.caption) $F_LB 10 $GREY 1.8 1 1 1 | Out-Null
    }
    Footer $Slide $WHITE
}

function L-Statement {
    param($Slide,$C)
    Wordmark $Slide $INK
    Head $Slide ([string]$C.section) ([string]$C.title) ''
    Rect $Slide 'cleadbar' 48 282 116 18 $INK | Out-Null
    Txt $Slide 'cleadtxt' 48 282 116 18 ([string]$C.leadLabel) $F_LB 9 $WHITE 1.4 1 2 3 | Out-Null
    Txt $Slide 'clead' 48 312 330 76 ([string]$C.lead) $F_ZH_B 24 $INK 0 1.35 1 1 | Out-Null
    Txt $Slide 'cexp' 48 400 330 86 ([string]$C.explain) $F_ZH_R 14 $INK 0 1.5 1 1 | Out-Null
    Rect $Slide 'cpanel' 420 282 492 202 $WHITE | Out-Null
    Rect $Slide 'cpanelbar' 420 282 5 202 $GOLD | Out-Null
    Txt $Slide 'cptitle' 444 304 444 28 ([string]$C.panelTitle) $F_ZH_B 17 $INK 0 1 1 1 | Out-Null
    Rule $Slide 'cprule' 444 342 444 $RULE 1 | Out-Null
    $y = 356; $i = 0
    foreach ($r in $C.rows) {
        $i++
        Txt $Slide ('cpr'+$i) 444 $y 444 26 ([string]$r) $F_ZH_N 13 $INK | Out-Null
        $y += 30
        if ($i -lt $C.rows.Count) { Rule $Slide ('cpl'+$i) 444 ($y-8) 444 $RULE 1 | Out-Null }
    }
    if ($C.note) { Txt $Slide 'cnote' 48 486 620 20 ([string]$C.note) $F_LB 10 $GREY 1.4 1 1 1 | Out-Null }
    Footer $Slide $INK
}

function L-Data {
    param($Slide,$C)
    Wordmark $Slide $INK
    Head $Slide ([string]$C.section) ([string]$C.title) ([string]$C.kicker)
    Rule $Slide 'dgrid' 80 236 800 $RULE 1 | Out-Null
    Rule $Slide 'daxis' 80 400 800 $INK 1 | Out-Null
    $xs = @(135, 335, 535, 735)
    $i = 0
    foreach ($b in $C.bars) {
        $x = [double]$xs[$i]; $h = [double]$b.height
        $col = $INK
        if ($b.accent) { $col = $GOLD }
        Rect $Slide ('bar'+$i) $x (400 - $h) 100 $h $col | Out-Null
        Txt $Slide ('val'+$i) ($x-15) (400 - $h - 32) 130 30 ([string]$b.value) $F_EN 18 $INK 0 1 2 3 | Out-Null
        Txt $Slide ('cat'+$i) ($x-15) 410 130 24 ([string]$b.label) $F_ZH_N 12 $GREY 0 1 2 1 | Out-Null
        $i++
    }
    if ($C.readout) { Txt $Slide 'dread' 48 452 864 30 ([string]$C.readout) $F_ZH_N 15 $INK | Out-Null }
    if ($C.note)    { Txt $Slide 'dnote' 48 486 620 20 ([string]$C.note) $F_LB 10 $GREY 1.4 1 1 1 | Out-Null }
    Footer $Slide $INK
}

function L-Photos {
    param($Slide,$C)
    Wordmark $Slide $INK
    Head $Slide ([string]$C.section) ([string]$C.title) ([string]$C.kicker)
    $bw = 272.0; $bh = 172.0
    $n = $C.photos.Count
    $gap = 60.0
    if ($n -ge 3) { $gap = 44.0 }
    $total = $n * $bw + ($n - 1) * $gap
    $x0 = $Cx0 + ($CW - $total) / 2.0
    $i = 0
    foreach ($p in $C.photos) {
        $x = $x0 + $i * ($bw + $gap)
        Pic $Slide ('ph'+$i) ([string]$p.src) $x 286 $bw $bh | Out-Null
        Rule $Slide ('phr'+$i) $x 468 $bw $RULE 1 | Out-Null
        Txt $Slide ('phc'+$i) $x 474 ($bw - 30) 20 ([string]$p.caption) $F_LB 9.5 $GREY 1.3 1 1 1 | Out-Null
        $i++
    }
    Footer $Slide $INK
}

function L-StatementImage {
    param($Slide,$C)
    Wordmark $Slide $INK
    Head $Slide ([string]$C.section) ([string]$C.title) ''
    Rect $Slide 'sipanel' 480 280 432 208 $WHITE | Out-Null
    Rect $Slide 'sibar' 480 280 5 208 $GOLD | Out-Null
    Txt $Slide 'silead' 504 302 380 62 ([string]$C.lead) $F_ZH_B 20 $INK 0 1.3 1 1 | Out-Null
    Rule $Slide 'sirule' 504 378 380 $RULE 1 | Out-Null
    Txt $Slide 'sitext' 504 392 380 84 ([string]$C.explain) $F_ZH_R 13 $INK 0 1.5 1 1 | Out-Null
    if ($C.src) { Pic $Slide 'siimg' ([string]$C.src) 48 280 396 208 ([double]$C.cropBias) | Out-Null }
    if ($C.readout) { Txt $Slide 'siread' 48 500 864 24 ([string]$C.readout) $F_LB 10 $GREY 1.5 1 1 1 | Out-Null }
    Footer $Slide $INK
}

function L-StatementPhoto {
    param($Slide,$C)
    Wordmark $Slide $WHITE
    Txt $Slide 'spkick' 48 150 500 22 ([string]$C.kicker) $F_LB 10 $GREY2 3 1 1 1 | Out-Null
    Txt $Slide 'sptitle' 48 176 620 56 ([string]$C.title) $F_ZH_H 36 $WHITE 2 1 1 3 | Out-Null
    Rule $Slide 'sprule' 48 248 864 $WHITE 1 0.35 | Out-Null
    Pic $Slide 'spimg' ([string]$C.src) 48 284 864 196 ([double]$C.cropBias) | Out-Null
    if ($C.caption) { Txt $Slide 'spcap' 48 488 864 20 ([string]$C.caption) $F_LB 10 $GREY2 1.6 1 1 1 | Out-Null }
    Txt $Slide 'spnote' 48 508 600 20 ([string]$C.note) $F_LB 10 $GREY2 1.6 1 1 1 | Out-Null
    Footer $Slide $WHITE
}

function L-Closing {    param($Slide,$C)
    Wordmark $Slide $WHITE
    Txt $Slide 'ekick' 48 176 600 34 ([string]$C.kicker) $F_LB 15 $WHITE 4 1 1 3 | Out-Null
    Rect $Slide 'eband' 48 232 460 46 $GOLD | Out-Null
    Txt $Slide 'ebandtxt' 48 232 460 46 ([string]$C.band) $F_ZH_R 20 $INK 2 1 2 3 | Out-Null
    if ($C.mark) { Txt $Slide 'emark' 48 296 500 50 ([string]$C.mark) $F_EN 34 $WHITE 2 1 1 3 | Out-Null }
    Rule $Slide 'erule' 48 396 864 $WHITE 1 0.35 | Out-Null
    if ($C.note) { Txt $Slide 'enote' 48 412 864 24 ([string]$C.note) $F_LB 10 $GREY2 1.8 1 1 1 | Out-Null }
    Footer $Slide $WHITE
}

# ---------------- 主流程 ----------------
$cfg = Get-Content -LiteralPath $ContentPath -Raw -Encoding UTF8 | ConvertFrom-Json
$title = 'RHINE LAB'
if ($cfg.title) { $title = [string]$cfg.title }

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(0)
$pres.PageSetup.SlideWidth = $W
$pres.PageSetup.SlideHeight = $H
try {
    $n = 0
    foreach ($sd in $cfg.slides) {
        $n++
        $dark = ($sd.layout -eq 'dark' -or $sd.layout -eq 'closing' -or $sd.layout -eq 'statement-photo')
        $slide = Add-Slide ($(if ($dark) { $DARK } else { $IVORY }))
        switch ([string]$sd.layout) {
            'cover'            { L-Cover     $slide $sd }
            'agenda'           { L-Agenda    $slide $sd }
            'process'          { L-Process   $slide $sd }
            'dark'             { L-Dark      $slide $sd }
            'statement'        { L-Statement $slide $sd }
            'data'             { L-Data      $slide $sd }
            'photos'           { L-Photos    $slide $sd }
            'statement-image'  { L-StatementImage $slide $sd }
            'statement-photo'  { L-StatementPhoto $slide $sd }
            'closing'          { L-Closing   $slide $sd }
            default            { throw "未知版式: $($sd.layout)" }
        }
        Write-Output ("  [{0}/{1}] {2}" -f $n, $cfg.slides.Count, $sd.layout)    }

    try { $pres.BuiltInDocumentProperties.Item('Title').Value = $title } catch { Write-Output "title skipped: $($_.Exception.Message)" }
    if (Test-Path -LiteralPath $PptxOut) { Remove-Item -LiteralPath $PptxOut -Force }
    $pres.SaveAs($PptxOut, 24)   # ppSaveAsOpenXMLPresentation
    Write-Output "SAVED: $PptxOut"

    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $pres.Slides.Item($i).Export((Join-Path $PngDir ("m{0:d2}.png" -f $i)), 'PNG', 1600, 900)
    }
    $pres.Close()
} finally {
    try { $ppt.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch { }
}
Write-Output 'DONE'
Get-ChildItem $PngDir -Filter *.png | Select-Object Name,Length | Format-Table -AutoSize
