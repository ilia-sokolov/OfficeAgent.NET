<#
.SYNOPSIS
Opens every file of an OfficeAgent native corpus in desktop Word, PowerPoint or Excel and checks
what Office itself observes against the corpus manifest.

.DESCRIPTION
The corpus comes from NativeCorpusTests with OFFICEAGENT_NATIVE_CORPUS_OUTPUT set. For each case
the harness opens the output with repair disabled and alerts suppressed, checks the manifest's
expectations through the Office object model, saves a copy and reopens it, and, for visual cases,
exports a PDF. Word cases with revisions are also accepted and rejected in Word.

Repair detection is proved, not assumed: the corpus carries corrupt control files, and the run
fails unless the harness reports every one of them as refused or repaired.

Windows only; needs desktop Office. Writes results.json and results.md next to the manifest.

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/native/office_check.ps1 -Corpus C:\temp\corpus
#>
param(
    [Parameter(Mandatory = $true)][string]$Corpus,
    [string]$Only = '',
    [switch]$Trace,
    [int]$TimeoutSeconds = 180,
    # Internal: run one case or control in this process and write its result here.
    [string]$Worker = '',
    [string]$Id = ''
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

$manifest = Get-Content -Raw -Encoding UTF8 (Join-Path $Corpus 'manifest.json') | ConvertFrom-Json
$roundtrip = Join-Path $Corpus 'roundtrip'
$pdf = Join-Path $Corpus 'pdf'
New-Item -ItemType Directory -Force $roundtrip, $pdf | Out-Null

function Hash([string]$path) { (Get-FileHash -Algorithm SHA256 $path).Hash.ToLowerInvariant() }

$results = New-Object System.Collections.ArrayList
$script:checks = $null

function Step([string]$what) { if ($Trace) { Write-Host "    $((Get-Date).ToString('HH:mm:ss.fff')) $what" } }
function Check([string]$name, $expected, $actual, [bool]$pass) {
    Step "check $name"
    [void]$script:checks.Add([ordered]@{ name = $name; expected = $expected; actual = $actual; pass = $pass })
}
function Has($expect, [string]$key) { $null -ne $expect.PSObject.Properties[$key] }
function ContainsAll([string]$text, $items, [string]$label) {
    foreach ($s in $items) { Check "$label contains '$s'" $true ($text.Contains($s)) ($text.Contains($s)) }
}
function ExcludesAll([string]$text, $items, [string]$label) {
    foreach ($s in $items) { Check "$label excludes '$s'" $false ($text.Contains($s)) (-not $text.Contains($s)) }
}

# Excel writes a repair log to %TEMP% when it repairs a workbook on open.
function RepairLogs { @(Get-ChildItem -Path $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName) }

# ── Word ──────────────────────────────────────────────────────────────────────────────────

$wdAlertsNone = 0
$wdFormatDocumentDefault = 16
$wdExportFormatPDF = 17
$wdStatisticPages = 2

# Word's Open takes its arguments by reference, which PowerShell cannot pass positionally with
# gaps, so they go by name through late binding.
function Invoke-Named($target, [string]$method, [hashtable]$named) {
    $names = @($named.Keys)
    # Unwrapped: COM rejects a PowerShell-wrapped string with DISP_E_TYPEMISMATCH.
    $values = @($names | ForEach-Object { $v = $named[$_]; if ($v -is [psobject]) { $v.PSObject.BaseObject } else { $v } })
    [System.__ComObject].InvokeMember($method, [Reflection.BindingFlags]::InvokeMethod, $null, $target,
        [object[]]$values, $null, $null, [string[]]$names)
}

# The document gets a window, inside the hidden application: accept and reject need one.
function Open-Word($app, [string]$path, [bool]$readOnly) {
    Invoke-Named $app.Documents 'Open' ([ordered]@{
        FileName = $path; ConfirmConversions = $false; ReadOnly = $readOnly; AddToRecentFiles = $false
        Revert = $true; Visible = $true; OpenAndRepair = $false; NoEncodingDialog = $true })
}

function WordText($doc) {
    $parts = @($doc.Content.Text)
    foreach ($s in $doc.Sections) { foreach ($i in 1..3) {
        $parts += $s.Headers.Item($i).Range.Text; $parts += $s.Footers.Item($i).Range.Text } }
    ($parts -join "`n") -replace "`r", "`n"
}

function Check-Word($app, $case, [string]$path) {
    $e = $case.expect
    Step 'open'
    $doc = Open-Word $app $path $true
    try {
        Step 'text'
        $script:observed['wordCompatibilityMode'] = $doc.CompatibilityMode
        $text = WordText $doc
        $revisions = $doc.Revisions.Count
        $tracked = (Has $e 'revisionsAtLeast') -or (Has $e 'revisionAuthors')
        if (Has $e 'revisions') { Check 'revisions' $e.revisions $revisions ($revisions -eq $e.revisions) }
        if (Has $e 'revisionsAtLeast') { Check 'revisions at least' $e.revisionsAtLeast $revisions ($revisions -ge $e.revisionsAtLeast) }
        if (Has $e 'revisionAuthors') {
            $authors = @($doc.Revisions | ForEach-Object { $_.Author } | Sort-Object -Unique)
            Check 'revision authors' ($e.revisionAuthors -join ', ') ($authors -join ', ') ((@($e.revisionAuthors) -join '|') -eq ($authors -join '|'))
        }
        if ((Has $e 'textContains') -and -not $tracked) { ContainsAll $text $e.textContains 'text' }
        if ((Has $e 'textExcludes') -and -not $tracked) { ExcludesAll $text $e.textExcludes 'text' }
        if (Has $e 'commentsAtLeast') { Check 'comments at least' $e.commentsAtLeast $doc.Comments.Count ($doc.Comments.Count -ge $e.commentsAtLeast) }
        if (Has $e 'footnotes') { Check 'footnotes' $e.footnotes $doc.Footnotes.Count ($doc.Footnotes.Count -eq $e.footnotes) }
        if (Has $e 'footnoteText') { $t = $doc.Footnotes.Item(1).Range.Text.Trim(); Check 'footnote text' $e.footnoteText $t ($t -eq $e.footnoteText) }
        if (Has $e 'endnotes') { Check 'endnotes' $e.endnotes $doc.Endnotes.Count ($doc.Endnotes.Count -eq $e.endnotes) }
        if (Has $e 'endnoteText') { $t = $doc.Endnotes.Item(1).Range.Text.Trim(); Check 'endnote text' $e.endnoteText $t ($t -eq $e.endnoteText) }
        if (Has $e 'tables') { Check 'tables' $e.tables $doc.Tables.Count ($doc.Tables.Count -eq $e.tables) }
        if (Has $e 'tableRows') { $n = $doc.Tables.Item(1).Rows.Count; Check 'table rows' $e.tableRows $n ($n -eq $e.tableRows) }
        if (Has $e 'tableColumns') { $n = $doc.Tables.Item(1).Columns.Count; Check 'table columns' $e.tableColumns $n ($n -eq $e.tableColumns) }
        if (Has $e 'inlineShapeSizePt') {
            $shape = $doc.InlineShapes.Item(1); $want = @($e.inlineShapeSizePt)
            $ok = [math]::Abs($shape.Width - $want[0]) -le 1 -and [math]::Abs($shape.Height - $want[1]) -le 1
            Check 'picture size in points' ($want -join ' x ') ('{0} x {1}' -f $shape.Width, $shape.Height) $ok
        }
        if (Has $e 'inlineShapes') { Check 'inline pictures' $e.inlineShapes $doc.InlineShapes.Count ($doc.InlineShapes.Count -eq $e.inlineShapes) }
        if (Has $e 'contentControlsAtLeast') { Check 'content controls at least' $e.contentControlsAtLeast $doc.ContentControls.Count ($doc.ContentControls.Count -ge $e.contentControlsAtLeast) }
        if (Has $e 'sectionsAtLeast') { Check 'sections at least' $e.sectionsAtLeast $doc.Sections.Count ($doc.Sections.Count -ge $e.sectionsAtLeast) }
        if (Has $e 'pagesAtLeast') { $n = $doc.ComputeStatistics($wdStatisticPages); Check 'pages at least' $e.pagesAtLeast $n ($n -ge $e.pagesAtLeast) }
        if (Has $e 'landscape') { $o = $doc.Sections.Item(1).PageSetup.Orientation; Check 'landscape' $true ($o -eq 1) ($o -eq 1) }
        if (Has $e 'headerContains') { $t = $doc.Sections.Item(1).Headers.Item(1).Range.Text; Check "header contains '$($e.headerContains)'" $true $t.Contains($e.headerContains) $t.Contains($e.headerContains) }
        if (Has $e 'footerContains') { $t = $doc.Sections.Item(1).Footers.Item(1).Range.Text; Check "footer contains '$($e.footerContains)'" $true $t.Contains($e.footerContains) $t.Contains($e.footerContains) }
        if (Has $e 'title') {
            # Document properties are IDispatch-only; PowerShell's binder cannot call them directly.
            $property = [System.__ComObject].InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $doc.BuiltInDocumentProperties, @('Title'))
            $t = [string][System.__ComObject].InvokeMember('Value', [Reflection.BindingFlags]::GetProperty, $null, $property, $null)
            Check 'title property' $e.title $t ($t -eq $e.title)
        }
        if (Has $e 'styleExists') {
            $found = $false; try { $null = $doc.Styles.Item($e.styleExists); $found = $true } catch { }
            Check "style '$($e.styleExists)' exists" $true $found $found
        }
        if (Has $e 'headerPictureBehindText') {
            # The engine anchors a page-sized picture behind text in the header, as Word's own
            # designed templates do, rather than using w:background.
            $behind = @($doc.Sections.Item(1).Headers.Item(1).Shapes | Where-Object { $_.Type -eq $msoPicture -and $_.WrapFormat.Type -eq 5 }).Count
            Check 'header holds a picture behind text' 1 $behind ($behind -ge 1)
        }
        if (Has $e 'wordCompatibilityMode') { $m = $doc.CompatibilityMode; Check 'Word compatibility mode' $e.wordCompatibilityMode $m ($m -eq $e.wordCompatibilityMode) }
        if (Has $e 'wordTextContains') { ContainsAll $text $e.wordTextContains 'text' }
        if (Has $e 'wordTextSameAsInput') {
            $source = Open-Word $app (Join-Path (Join-Path $Corpus 'inputs') $case.file) $true
            try { $was = WordText $source } finally { $source.Close($false) }
            Check 'Word reads the same text as in the input' $was $text ($was -eq $text)
        }

        if ((Has $e 'acceptContains') -or (Has $e 'acceptExcludes') -or (Has $e 'acceptTableColumns')) {
            $doc.Revisions.AcceptAll()
            $t = WordText $doc
            Check 'revisions after accept all' 0 $doc.Revisions.Count ($doc.Revisions.Count -eq 0)
            if (Has $e 'acceptContains') { ContainsAll $t $e.acceptContains 'accepted text' }
            if (Has $e 'acceptExcludes') { ExcludesAll $t $e.acceptExcludes 'accepted text' }
            if (Has $e 'acceptTableColumns') { $n = $doc.Tables.Item(1).Columns.Count; Check 'table columns after accept all' $e.acceptTableColumns $n ($n -eq $e.acceptTableColumns) }
        }
    } finally { $doc.Close($false) }

    if ((Has $e 'rejectContains') -or (Has $e 'rejectExcludes') -or (Has $e 'rejectTableColumns')) {
        $doc = Open-Word $app $path $true
        try {
            $doc.Revisions.RejectAll()
            $t = WordText $doc
            Check 'revisions after reject all' 0 $doc.Revisions.Count ($doc.Revisions.Count -eq 0)
            if (Has $e 'rejectContains') { ContainsAll $t $e.rejectContains 'rejected text' }
            if (Has $e 'rejectExcludes') { ExcludesAll $t $e.rejectExcludes 'rejected text' }
            if (Has $e 'rejectTableColumns') { $n = $doc.Tables.Item(1).Columns.Count; Check 'table columns after reject all' $e.rejectTableColumns $n ($n -eq $e.rejectTableColumns) }
        } finally { $doc.Close($false) }
    }

    # Save as Word itself would, reopen the saved file, and export a PDF for visual review.
    $saved = Join-Path $roundtrip $case.file
    $doc = Open-Word $app $path $true
    try {
        Step 'save as'
        # CompatibilityMode 0 keeps the file's mode; positional SaveAs2 blocks on an upgrade prompt.
        Invoke-Named $doc 'SaveAs2' ([ordered]@{ FileName = $saved; FileFormat = $wdFormatDocumentDefault; AddToRecentFiles = $false; CompatibilityMode = 0 }) | Out-Null
        Step 'export'
        # Named, like SaveAs2: the positional call can block the automation session.
        if ($case.visual) { Invoke-Named $doc 'ExportAsFixedFormat' ([ordered]@{ OutputFileName = (Join-Path $pdf ($case.id + '.pdf')); ExportFormat = $wdExportFormatPDF; OpenAfterExport = $false }) | Out-Null }
    } finally { $doc.Close($false) }
    if ($case.visual) {
        # The input too, so a reviewer compares before and after rather than judging alone.
        $source = Open-Word $app (Join-Path (Join-Path $Corpus 'inputs') $case.file) $true
        try { Invoke-Named $source 'ExportAsFixedFormat' ([ordered]@{ OutputFileName = (Join-Path $pdf ($case.id + '.input.pdf')); ExportFormat = $wdExportFormatPDF; OpenAfterExport = $false }) | Out-Null }
        finally { $source.Close($false) }
    }
    $again = Open-Word $app $saved $true
    try { Check 'Word-saved copy reopens with the same revision count' $revisions $again.Revisions.Count ($again.Revisions.Count -eq $revisions) }
    finally { $again.Close($false) }
}

# ── PowerPoint ────────────────────────────────────────────────────────────────────────────

$ppSaveAsOpenXMLPresentation = 24
$ppSaveAsPDF = 32
$msoPicture = 13
$msoMedia = 16
$msoFillPicture = 6

function Open-Deck($app, [string]$path) {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $app.Presentations.Open($path, -1, 0, 0)
}

function DeckText($deck) {
    $parts = @()
    foreach ($slide in $deck.Slides) { foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $parts += $shape.TextFrame.TextRange.Text }
        if ($shape.HasTable) { foreach ($r in 1..$shape.Table.Rows.Count) { foreach ($c in 1..$shape.Table.Columns.Count) {
            $parts += $shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text } } }
    } }
    ($parts -join "`n") -replace "`r", "`n"
}

function Shapes($deck) { foreach ($slide in $deck.Slides) { foreach ($shape in $slide.Shapes) { $shape } } }

function Check-Deck($app, $case, [string]$path) {
    $e = $case.expect
    $deck = Open-Deck $app $path
    try {
        $text = DeckText $deck
        if (Has $e 'slides') { Check 'slides' $e.slides $deck.Slides.Count ($deck.Slides.Count -eq $e.slides) }
        if (Has $e 'textContains') { ContainsAll $text $e.textContains 'text' }
        if (Has $e 'textExcludes') { ExcludesAll $text $e.textExcludes 'text' }
        if (Has $e 'firstSlideContains') {
            $t = (@($deck.Slides.Item(1).Shapes | Where-Object { $_.HasTextFrame -and $_.TextFrame.HasText } | ForEach-Object { $_.TextFrame.TextRange.Text }) -join ' ')
            Check "first slide contains '$($e.firstSlideContains)'" $true $t.Contains($e.firstSlideContains) $t.Contains($e.firstSlideContains)
        }
        $shapes = @(Shapes $deck)
        if (Has $e 'tables') { $n = @($shapes | Where-Object { $_.HasTable }).Count; Check 'tables' $e.tables $n ($n -eq $e.tables) }
        if (Has $e 'pictures') { $n = @($shapes | Where-Object { $_.Type -eq $msoPicture }).Count; Check 'pictures' $e.pictures $n ($n -eq $e.pictures) }
        if (Has $e 'media') { $n = @($shapes | Where-Object { $_.Type -eq $msoMedia }).Count; Check 'media objects' $e.media $n ($n -eq $e.media) }
        if (Has $e 'charts') { $n = @($shapes | Where-Object { $_.HasChart }).Count; Check 'charts' $e.charts $n ($n -eq $e.charts) }
        if (Has $e 'chartEditable') {
            $chart = @($shapes | Where-Object { $_.HasChart })[0].Chart
            $editable = $false
            try {
                $chart.ChartData.Activate()
                $book = $chart.ChartData.Workbook
                $editable = $null -ne $book.Worksheets.Item(1).UsedRange
                $book.Application.Quit()
            } catch { $editable = $false }
            Check 'chart data opens for editing' $true $editable $editable
        }
        if (Has $e 'chartSeriesValues') {
            $chart = @($shapes | Where-Object { $_.HasChart })[0].Chart
            $values = @($chart.SeriesCollection(1).Values | ForEach-Object { [double]$_ })
            Check 'chart series values' ($e.chartSeriesValues -join ', ') ($values -join ', ') ((@($e.chartSeriesValues) -join '|') -eq ($values -join '|'))
        }
        if (Has $e 'commentsAtLeast') { $n = 0; foreach ($s in $deck.Slides) { $n += $s.Comments.Count }; Check 'comments at least' $e.commentsAtLeast $n ($n -ge $e.commentsAtLeast) }
        if (Has $e 'sectionsAtLeast') { $n = $deck.SectionProperties.Count; Check 'sections at least' $e.sectionsAtLeast $n ($n -ge $e.sectionsAtLeast) }
        if (Has $e 'sectionNames') {
            $names = @(); foreach ($i in 1..$deck.SectionProperties.Count) { $names += $deck.SectionProperties.Name($i) }
            foreach ($name in $e.sectionNames) { Check "section '$name'" $true ($names -contains $name) ($names -contains $name) }
        }
        if (Has $e 'transitionsOnAll') {
            $all = $true; foreach ($s in $deck.Slides) { if ($s.SlideShowTransition.EntryEffect -eq 0) { $all = $false } }
            Check 'every slide has a transition' $true $all $all
        }
        if (Has $e 'animationsAtLeast') { $n = 0; foreach ($s in $deck.Slides) { $n += $s.TimeLine.MainSequence.Count }; Check 'animations at least' $e.animationsAtLeast $n ($n -ge $e.animationsAtLeast) }
        if (Has $e 'footerContains') {
            $t = $deck.Slides.Item(1).HeadersFooters.Footer.Text
            Check "footer contains '$($e.footerContains)'" $true $t.Contains($e.footerContains) $t.Contains($e.footerContains)
        }
        if (Has $e 'backgroundPicture') {
            $s = $deck.Slides.Item(1)
            $ok = (-not $s.FollowMasterBackground) -and ($s.Background.Fill.Type -eq $msoFillPicture)
            Check 'slide background is a picture' $true $ok $ok
        }
        $slides = $deck.Slides.Count
        $saved = Join-Path $roundtrip $case.file
        $deck.SaveCopyAs($saved, $ppSaveAsOpenXMLPresentation)
        if ($case.visual) { $deck.SaveCopyAs((Join-Path $pdf ($case.id + '.pdf')), $ppSaveAsPDF) }
    } finally { $deck.Close() }
    if ($case.visual) {
        $source = Open-Deck $app (Join-Path (Join-Path $Corpus 'inputs') $case.file)
        try { $source.SaveCopyAs((Join-Path $pdf ($case.id + '.input.pdf')), $ppSaveAsPDF) } finally { $source.Close() }
    }
    $again = Open-Deck $app $saved
    try { Check 'PowerPoint-saved copy reopens with the same slide count' $slides $again.Slides.Count ($again.Slides.Count -eq $slides) }
    finally { $again.Close() }
}

# ── Excel ─────────────────────────────────────────────────────────────────────────────────

$xlOpenXMLWorkbook = 51
$xlTypePDF = 0

function Open-Book($app, [string]$path) {
    # CorruptLoad = xlNormalLoad: no repair or data extraction is requested.
    Invoke-Named $app.Workbooks 'Open' ([ordered]@{
        Filename = $path; UpdateLinks = 0; ReadOnly = $true; IgnoreReadOnlyRecommended = $true
        Notify = $false; AddToMru = $false; CorruptLoad = 0 })
}

function Check-Book($app, $case, [string]$path) {
    $e = $case.expect
    if (Has $e 'tableRowsGrewBy') {
        # Read first: Excel will not open two workbooks with the same file name at once.
        $source = Open-Book $app (Join-Path (Join-Path $Corpus 'inputs') $case.file)
        try { $inputRows = $source.Worksheets.Item(1).ListObjects.Item(1).ListRows.Count } finally { $source.Close($false) }
    }
    $before = RepairLogs
    $book = Open-Book $app $path
    try {
        $sheet = $book.Worksheets.Item(1)
        $logs = @(RepairLogs | Where-Object { $before -notcontains $_ })
        Check 'no Excel repair log' 0 $logs.Count ($logs.Count -eq 0)
        if (Has $e 'recalculate') { $app.CalculateFull() }
        if (Has $e 'cells') {
            foreach ($cell in $e.cells) {
                $range = $sheet.Range($cell.address)
                if ($null -ne $cell.PSObject.Properties['value']) {
                    $v = [string]$range.Value2
                    Check "$($cell.address) value" $cell.value $v ($v -eq $cell.value)
                }
                if ($null -ne $cell.PSObject.Properties['formula']) {
                    $f = [string]$range.Formula
                    Check "$($cell.address) formula" $cell.formula $f ($f -eq $cell.formula)
                }
            }
        }
        $text = (@($sheet.UsedRange.Value2) | ForEach-Object { [string]$_ }) -join "`n"
        if (Has $e 'textContains') { ContainsAll $text $e.textContains 'cells' }
        if (Has $e 'commentsAtLeast') {
            $n = $sheet.Comments.Count + $sheet.CommentsThreaded.Count
            Check 'comments at least' $e.commentsAtLeast $n ($n -ge $e.commentsAtLeast)
        }
        if (Has $e 'tables') { $n = 0; foreach ($ws in $book.Worksheets) { $n += $ws.ListObjects.Count }; Check 'tables' $e.tables $n ($n -eq $e.tables) }
        if (Has $e 'tableRowsGrewBy') {
            $now = $sheet.ListObjects.Item(1).ListRows.Count
            Check 'table data rows grew by' $e.tableRowsGrewBy ($now - $inputRows) (($now - $inputRows) -eq $e.tableRowsGrewBy)
        }
        $saved = Join-Path $roundtrip $case.file
        $book.SaveCopyAs($saved)
        if ($case.visual) { $book.ExportAsFixedFormat($xlTypePDF, (Join-Path $pdf ($case.id + '.pdf'))) }
    } finally { $book.Close($false) }
    if ($case.visual) {
        $source = Open-Book $app (Join-Path (Join-Path $Corpus 'inputs') $case.file)
        try { $source.ExportAsFixedFormat($xlTypePDF, (Join-Path $pdf ($case.id + '.input.pdf'))) } finally { $source.Close($false) }
    }
    $again = Open-Book $app $saved
    try { Check 'Excel-saved copy reopens' $true $true $true } finally { $again.Close($false) }
}

# ── run ───────────────────────────────────────────────────────────────────────────────────

function New-App([string]$prog) {
    $app = New-Object -ComObject $prog
    switch ($prog) {
        'Word.Application' { $app.Visible = $false; $app.DisplayAlerts = $wdAlertsNone }
        'PowerPoint.Application' { $app.DisplayAlerts = 1 }  # ppAlertsNone
        'Excel.Application' { $app.Visible = $false; $app.DisplayAlerts = $false; $app.AskToUpdateLinks = $false }
    }
    $app
}

$progs = @{ word = 'Word.Application'; powerpoint = 'PowerPoint.Application'; excel = 'Excel.Application' }
$apps = @{}
function App([string]$format) {
    if (-not $apps.ContainsKey($format)) { $apps[$format] = New-App $progs[$format] }
    $apps[$format]
}

function Run-Case($case, [string]$path) {
    $script:checks = New-Object System.Collections.ArrayList
    $script:observed = [ordered]@{}
    $opened = $true; $failure = $null
    try {
        switch ($case.format) {
            'word' { Check-Word (App 'word') $case $path }
            'powerpoint' { Check-Deck (App 'powerpoint') $case $path }
            'excel' { Check-Book (App 'excel') $case $path }
        }
    } catch { $opened = $false; $failure = $_.Exception.Message }
    $pass = $opened -and (@($script:checks | Where-Object { -not $_.pass }).Count -eq 0)
    [ordered]@{
        id = $case.id; format = $case.format; family = $case.family; description = $case.description
        sha256 = (Hash $path); manifestSha256 = $case.outputSha256
        openedCleanly = $opened; error = $failure; pass = $pass; observed = $script:observed; checks = @($script:checks)
    }
}

function Control-Case([IO.FileInfo]$file) {
    $format = @{ '.docx' = 'word'; '.pptx' = 'powerpoint'; '.xlsx' = 'excel' }[$file.Extension]
    [pscustomobject]@{ id = $file.BaseName; format = $format; family = 'control'; description = 'corrupt control'
        file = $file.Name; visual = $false; outputSha256 = ''; expect = [pscustomobject]@{} }
}

# Worker: one case in this process, so a blocking dialog cannot stall the whole run.
if ($Worker) {
    $control = Get-Item (Join-Path (Join-Path $Corpus 'controls') "$Id.*") -ErrorAction SilentlyContinue
    if ($control) { $case = Control-Case $control; $path = $control.FullName }
    else { $case = $manifest | Where-Object { $_.id -eq $Id }; $path = Join-Path (Join-Path $Corpus 'outputs') $case.file }
    try { $r = Run-Case $case $path }
    finally { foreach ($app in $apps.Values) { try { $app.Quit() } catch { }; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
    $r | ConvertTo-Json -Depth 8 | Out-File -Encoding utf8 $Worker
    exit 0
}

$images = @{ word = 'WINWORD'; powerpoint = 'POWERPNT'; excel = 'EXCEL' }
function Run-Isolated($case) {
    $out = Join-Path $env:TEMP ("oa-native-" + [guid]::NewGuid().ToString('N') + '.json')
    $started = Get-Date
    $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Corpus', $Corpus, '-Worker', $out, '-Id', $case.id)
    if ($Trace) { $args += '-Trace' }
    $process = Start-Process powershell -ArgumentList $args -PassThru -NoNewWindow
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        # Hung, almost always on a modal dialog: kill the worker and the Office process it drove.
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Get-Process $images[$case.format] -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $started } | Stop-Process -Force
        return [ordered]@{ id = $case.id; format = $case.format; family = $case.family; description = $case.description
            openedCleanly = $false; error = "no answer within $TimeoutSeconds s (killed; a modal dialog is the usual cause)"
            pass = $false; hung = $true; observed = @{}; checks = @() }
    }
    if (-not (Test-Path $out)) {
        return [ordered]@{ id = $case.id; format = $case.format; family = $case.family; description = $case.description
            openedCleanly = $false; error = "worker exited $($process.ExitCode) without a result"; pass = $false; observed = @{}; checks = @() }
    }
    $r = Get-Content -Raw -Encoding UTF8 $out | ConvertFrom-Json
    Remove-Item $out
    $r
}

foreach ($case in $manifest) {
    if ($Only -and $case.id -notlike $Only) { continue }
    $r = Run-Isolated $case
    if ($r.pass -and $r.sha256 -ne $case.outputSha256) { $r.pass = $false; $r.error = 'output hash differs from the manifest' }
    [void]$results.Add($r)
    Write-Host ("{0,-6} {1}" -f ($(if ($r.pass) { 'PASS' } else { 'FAIL' })), $case.id)
}

# Negative controls: each must be refused, and it counts as caught only when the same
# application opened a valid case in this run; otherwise a harness that cannot open anything
# would "catch" every control.
$controls = New-Object System.Collections.ArrayList
foreach ($file in Get-ChildItem (Join-Path $Corpus 'controls')) {
    $case = Control-Case $file
    $r = Run-Isolated $case
    $appWorks = @($results | Where-Object { $_.format -eq $case.format -and $_.openedCleanly }).Count -gt 0
    $detected = (-not $r.pass) -and $appWorks
    [void]$controls.Add([ordered]@{ file = $file.Name; format = $case.format; detected = $detected; appOpenedValidCase = $appWorks; error = $r.error; checks = $r.checks })
    Write-Host ("{0,-6} {1}" -f ($(if ($detected) { 'CAUGHT' } else { 'MISSED' })), $file.Name)
}

$word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ErrorAction SilentlyContinue
$os = Get-CimInstance Win32_OperatingSystem
$environment = [ordered]@{
    officeBuild = $(if ($word) { $word.VersionInfo.ProductVersion } else { 'unknown' })
    os = "$($os.Caption) $($os.Version)"
    runUtc = (Get-Date).ToUniversalTime().ToString('o')
    manifestSha256 = (Hash (Join-Path $Corpus 'manifest.json'))
}
$report = [ordered]@{ environment = $environment; cases = @($results); controls = @($controls) }
$report | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 (Join-Path $Corpus 'results.json')

$failed = @($results | Where-Object { -not $_.pass })
$missed = @($controls | Where-Object { -not $_.detected })
Write-Host ""
Write-Host "Office $($environment.officeBuild) on $($environment.os)"
Write-Host "$($results.Count - $failed.Count)/$($results.Count) cases pass; $($controls.Count - $missed.Count)/$($controls.Count) corrupt controls caught"
if ($failed.Count -gt 0 -or $missed.Count -gt 0) { exit 1 }
