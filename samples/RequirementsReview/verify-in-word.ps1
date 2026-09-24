<#
.SYNOPSIS
  Opens a reviewed copy in desktop Word (read-only, invisible), reports what Word itself sees,
  and closes without saving. This is native verification; the test suite only checks the
  saved package structure.

.EXAMPLE
  ./verify-in-word.ps1 -Path ./review-out/method-statement.reviewed.docx
  ./verify-in-word.ps1 -Path ./review-out/method-statement.reviewed.docx -Pdf ./review-out/method-statement.reviewed.pdf
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$Pdf
)

$ErrorActionPreference = 'Stop'
$full = (Resolve-Path $Path).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
  $doc = $word.Documents.Open($full, $false, $true, $false)
  try {
    $revisions = @()
    foreach ($r in $doc.Revisions) {
      $kind = switch ($r.Type) { 1 { 'insert' } 2 { 'delete' } default { "type$($r.Type)" } }
      $revisions += [pscustomobject]@{ Author = $r.Author; Kind = $kind; Text = $r.Range.Text }
    }
    $comments = @()
    foreach ($c in $doc.Comments) {
      $comments += [pscustomobject]@{ Author = $c.Author; Anchor = $c.Scope.Text; Text = $c.Range.Text }
    }

    [pscustomobject]@{
      WordVersion = $word.Version
      Build       = $word.Build
      File        = [IO.Path]::GetFileName($full)
      Paragraphs  = $doc.Paragraphs.Count
      Tables      = $doc.Tables.Count
      TableRows   = if ($doc.Tables.Count -gt 0) { $doc.Tables.Item(1).Rows.Count } else { 0 }
      Revisions   = $revisions.Count
      Comments    = $comments.Count
    } | Format-List

    'Revisions:'; $revisions | Format-Table -AutoSize -Wrap
    'Comments:';  $comments  | Format-Table -AutoSize -Wrap

    if ($Pdf) {
      # 17 = wdExportFormatPDF, 7 = wdExportDocumentWithMarkup: what a reviewer sees with markup on.
      $doc.ExportAsFixedFormat([IO.Path]::GetFullPath($Pdf), 17, $false, 0, 0, 1, 1, 7)
      "Exported with markup: $Pdf"
    }
  }
  finally {
    $doc.Close([ref]0)  # wdDoNotSaveChanges
  }
}
finally {
  $word.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
