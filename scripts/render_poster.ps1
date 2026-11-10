[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [ValidateRange(300, 10000)][int]$Width = 2266,
    [switch]$Jpg,
    [switch]$Force
)

# PowerPoint is the renderer; the supplied presentation is never saved or edited.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

$source = (Resolve-Path -LiteralPath $InputPptx).ProviderPath
if ([IO.Path]::GetExtension($source) -ine '.pptx') { throw 'InputPptx must be a .pptx file.' }
$destination = [IO.Path]::GetFullPath($OutputDir)
if ([IO.File]::Exists($destination)) { throw 'OutputDir points to an existing file.' }
[void][IO.Directory]::CreateDirectory($destination)
$outputs = @('poster.png', 'poster.pdf', 'qa.json')
if ($Jpg) { $outputs += 'poster.jpg' }
foreach ($name in $outputs) {
    $target = Join-Path $destination $name
    if ($source -ieq $target) { throw 'An output would replace the input.' }
    if ((Test-Path -LiteralPath $target) -and -not $Force) {
        throw "Output already exists: $target. Use a new OutputDir or explicitly pass -Force."
    }
}

$scratch = Join-Path $destination ('.render-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
$copy = Join-Path $scratch 'input.pptx'
Copy-Item -LiteralPath $source -Destination $copy
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$app = $null
$presentations = $null
$deck = $null
$slides = $null
$slide = $null
$shapes = $null
$pageSetup = $null
$qaShapes = [Collections.Generic.List[object]]::new()

try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentations = $app.Presentations
    # ReadOnly = true, Untitled = false, WithWindow = false.
    # A unique scratch copy prevents reusing or closing an already-open user file.
    $deck = $presentations.Open($copy, -1, 0, 0)
    $slides = $deck.Slides
    if ($slides.Count -ne 1) { throw "Expected one poster slide; found $($slides.Count)." }
    $pageSetup = $deck.PageSetup
    $slideWidth = [double]$pageSetup.SlideWidth
    $slideHeight = [double]$pageSetup.SlideHeight
    $height = [int][Math]::Round($Width * $slideHeight / $slideWidth)
    $slide = $slides.Item(1)
    $shapes = $slide.Shapes
    for ($i = 1; $i -le $shapes.Count; $i++) {
        $shape = $null
        $frame = $null
        $range = $null
        try {
            $shape = $shapes.Item($i)
            if ($shape.HasTextFrame -ne -1) { continue }
            $frame = $shape.TextFrame2
            if ($frame.HasText -ne -1) { continue }
            $range = $frame.TextRange
            $text = [string]$range.Text
            $boxLeft = [double]$shape.Left
            $boxTop = [double]$shape.Top
            $boxWidth = [double]$shape.Width
            $boxHeight = [double]$shape.Height
            $boundLeft = [double]$range.BoundLeft
            $boundTop = [double]$range.BoundTop
            $boundWidth = [double]$range.BoundWidth
            $boundHeight = [double]$range.BoundHeight
            $fonts = [Collections.Generic.List[object]]::new()
            $runs = $null
            try {
                $runs = $range.Runs()
                for ($j = 1; $j -le $runs.Count; $j++) {
                    $run = $null
                    $font = $null
                    try {
                        $run = $range.Runs($j, 1)
                        $font = $run.Font
                        $fonts.Add([ordered]@{
                            text = [string]$run.Text
                            latin = [string]$font.Name
                            east_asian = [string]$font.NameFarEast
                            size_pt = [double]$font.Size
                        })
                    } finally {
                        Release-ComObject $font
                        Release-ComObject $run
                    }
                }
            } finally { Release-ComObject $runs }
            # Small tolerance avoids rounding and glyph-ink overhang false positives.
            $tolerance = 1.5
            $overflow = [ordered]@{
                left_pt = [Math]::Max([double]0, $boxLeft - $boundLeft)
                top_pt = [Math]::Max([double]0, $boxTop - $boundTop)
                right_pt = [Math]::Max([double]0, $boundLeft + $boundWidth - $boxLeft - $boxWidth)
                bottom_pt = [Math]::Max([double]0, $boundTop + $boundHeight - $boxTop - $boxHeight)
            }
            $possibleOverflow = ($overflow.left_pt -gt $tolerance -or $overflow.top_pt -gt $tolerance -or $overflow.right_pt -gt $tolerance -or $overflow.bottom_pt -gt $tolerance)
            $qaShapes.Add([ordered]@{
                id = [int]$shape.Id
                name = [string]$shape.Name
                text = $text
                box_pt = [ordered]@{left=$boxLeft;top=$boxTop;width=$boxWidth;height=$boxHeight}
                text_bounds_pt = [ordered]@{left=$boundLeft;top=$boundTop;width=$boundWidth;height=$boundHeight}
                margins_pt = [ordered]@{left=[double]$frame.MarginLeft;top=[double]$frame.MarginTop;right=[double]$frame.MarginRight;bottom=[double]$frame.MarginBottom}
                overflow_pt = $overflow
                possible_overflow = $possibleOverflow
                fonts = @($fonts.ToArray())
            })
        } finally {
            Release-ComObject $range
            Release-ComObject $frame
            Release-ComObject $shape
        }
    }
    # Render to temporary outputs before promoting the complete set.
    $slide.Export((Join-Path $scratch 'poster.png'), 'PNG', $Width, $height)
    if ($Jpg) { $slide.Export((Join-Path $scratch 'poster.jpg'), 'JPG', $Width, $height) }
    $deck.SaveCopyAs((Join-Path $scratch 'poster.pdf'), 32)
    $qa = [ordered]@{
        source_pptx = $source
        source_sha256 = $sourceHash
        renderer = 'Microsoft PowerPoint COM'
        renderer_version = [string]$app.Version
        slide_count = [int]$slides.Count
        slide_size_pt = [ordered]@{width=$slideWidth;height=$slideHeight}
        png_size_px = [ordered]@{width=$Width;height=$height}
        font_note = 'Font names are requested PowerPoint run fonts; installed-font substitution must be checked visually.'
        overflow_note = 'Bounds are diagnostic. Inspect the PNG for overlap, clipping, missing glyphs and font substitution.'
        possible_overflow_shape_ids = @($qaShapes | Where-Object { $_.possible_overflow } | ForEach-Object { $_.id })
        shapes = @($qaShapes.ToArray())
    }
    $qa | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $scratch 'qa.json') -Encoding UTF8
    foreach ($name in $outputs) {
        $rendered = Join-Path $scratch $name
        if (-not (Test-Path -LiteralPath $rendered) -or (Get-Item -LiteralPath $rendered).Length -eq 0) {
            throw "Renderer did not produce a non-empty $name."
        }
    }
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'The input changed during rendering; outputs have not been promoted.'
    }
    foreach ($name in $outputs) {
        Move-Item -LiteralPath (Join-Path $scratch $name) -Destination (Join-Path $destination $name) -Force:$Force
    }
    [ordered]@{
        output_dir = $destination
        png = (Join-Path $destination 'poster.png')
        pdf = (Join-Path $destination 'poster.pdf')
        qa = (Join-Path $destination 'qa.json')
        width = $Width
        height = $height
        possible_overflow_shape_ids = $qa.possible_overflow_shape_ids
    } | ConvertTo-Json -Depth 4
} finally {
    Release-ComObject $shapes
    Release-ComObject $slide
    Release-ComObject $slides
    Release-ComObject $pageSetup
    if ($null -ne $deck) {
        try { $deck.Close() } finally { Release-ComObject $deck }
    }
    Release-ComObject $presentations
    # Never call Application.Quit(): it could terminate the user's PowerPoint session.
    Release-ComObject $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # Delete only this exact, verified scratch directory owned by this invocation.
    $resolvedScratch = [IO.Path]::GetFullPath($scratch)
    $expectedParent = [IO.Path]::GetFullPath($destination).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($resolvedScratch) -ieq $expectedParent -and [IO.Path]::GetFileName($resolvedScratch) -match '^\.render-[0-9a-f]{32}$') {
        if (Test-Path -LiteralPath $resolvedScratch) { Remove-Item -LiteralPath $resolvedScratch -Recurse -Force }
    }
}
