param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][string]$Output,[switch]$ShowMarks)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $Output | Out-Null
$word=$null;$doc=$null
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0
    $doc=$word.Documents.Open($Path,$false,$true)
    $doc.ActiveWindow.View.Type=3
    $doc.ActiveWindow.View.ShowAll=$ShowMarks.IsPresent
    $pages=$doc.ActiveWindow.ActivePane.Pages
    Write-Output ('PAGES '+$pages.Count)
    for($i=1;$i -le $pages.Count;$i++) {
        $bits=[byte[]]$pages.Item($i).EnhMetaFileBits
        $stream=New-Object IO.MemoryStream(,$bits)
        $mf=[Drawing.Image]::FromStream($stream)
        $image=New-Object Drawing.Bitmap 1240,1754
        $g=[Drawing.Graphics]::FromImage($image)
        try {
            $g.Clear([Drawing.Color]::White)
            $g.DrawImage($mf,0,0,$image.Width,$image.Height)
            $image.Save((Join-Path $Output ('page-'+$i+'.png')),[Drawing.Imaging.ImageFormat]::Png)
        } finally {$g.Dispose();$image.Dispose();$mf.Dispose();$stream.Dispose()}
    }
} finally {if($doc){$doc.Close(0)};if($word){$word.Quit()}}
