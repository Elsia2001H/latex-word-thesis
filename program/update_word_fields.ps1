param([Parameter(Mandatory=$true)][string]$Path, [switch]$PageNumbersOnly)
$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($resolved, $false, $false)
    $document.Repaginate()
    if (-not $PageNumbersOnly) {
        $document.Fields.Update() | Out-Null
        foreach ($toc in $document.TablesOfContents) { $toc.Update() }
    }
    $document.Repaginate()
    foreach ($section in $document.Sections) {
        foreach ($header in $section.Headers) {
            if ($header.Exists) { $header.Range.Fields.Update() | Out-Null }
        }
    }
    foreach ($toc in $document.TablesOfContents) { $toc.UpdatePageNumbers() }
    # Bold the level-one title and tab leader, but not its page number.
    # Apply after all TOC updates because Word regenerates cached runs.
    $document.Styles.Item(-20).Font.Bold = -1 # wdStyleTOC1
    foreach ($toc in $document.TablesOfContents) {
        foreach ($paragraph in $toc.Range.Paragraphs) {
            if ($paragraph.Range.Style.NameLocal -ne $document.Styles.Item(-20).NameLocal) { continue }
            $range = $paragraph.Range.Duplicate
            $range.Font.Bold = -1
            $text = $range.Text
            $tabIndex = $text.LastIndexOf("`t")
            if ($tabIndex -ge 0) {
                $pageRange = $range.Duplicate
                $pageRange.Start = $range.Start + $tabIndex + 1
                $pageRange.End = $range.End - 1
                $pageRange.Font.Bold = 0
            }
        }
    }
    $document.Save()
} finally {
    if ($document) { $document.Close(0) }
    $word.Quit()
}
Write-Output "Word fields and TOC saved: $resolved"
