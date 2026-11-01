# Builds the encrypted fixtures with Office itself, so the detector is tested against the files
# Office really writes. The password is "fixture"; the documents hold one line of placeholder text.
# There is no Word fixture: Word blocked under automation here. The encrypted container is the same
# for every format, and the tests also run these bytes through the Word module.
param(
    [Parameter(Mandatory = $true)][ValidateSet('excel', 'powerpoint')][string]$Format,
    [string]$Out = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
$password = 'fixture'

switch ($Format) {
    'excel' {
        $excel = New-Object -ComObject Excel.Application
        try {
            $excel.DisplayAlerts = $false
            $book = $excel.Workbooks.Add()
            $book.Worksheets.Item(1).Range('A1').Value2 = 'Encrypted fixture'
            $book.Password = $password
            $book.SaveAs((Join-Path $Out 'encrypted.xlsx'), 51)
            $book.Close($false)
        } finally { $excel.Quit() }
    }
    'powerpoint' {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        try {
            $powerPoint.DisplayAlerts = 1
            $deck = $powerPoint.Presentations.Add(0)
            $deck.Slides.Add(1, 12) | Out-Null
            $deck.Password = $password
            $deck.SaveAs((Join-Path $Out 'encrypted.pptx'), 24)
            $deck.Close()
        } finally { $powerPoint.Quit() }
    }
}
