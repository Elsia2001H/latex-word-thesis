param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][int]$Expected,[string]$OwnerFile)
$ErrorActionPreference='Stop'
$word=$null;$doc=$null
try {
    $word=New-Object -ComObject Word.Application
    $word.Visible=$false;$word.DisplayAlerts=0
    $doc=$word.Documents.Open($Path,$false,$false)
    if(!$doc){throw 'Word returned no batch document'}
    if($OwnerFile){
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class MTWordOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
        [uint32]$workerPid=0
        [MTWordOwner]::GetWindowThreadProcessId([IntPtr]$doc.ActiveWindow.Hwnd,[ref]$workerPid) | Out-Null
        $workerProcess=Get-Process -Id $workerPid
        @{pid=$workerPid;startTicks=$workerProcess.StartTime.ToUniversalTime().Ticks;path=$Path} | ConvertTo-Json | Set-Content -LiteralPath $OwnerFile -Encoding UTF8
    }
    $doc.Content.Select()
    $word.GetType().InvokeMember('Run',[Reflection.BindingFlags]::InvokeMethod,$null,$word,
        [object[]]@('MathTypeCommands.UILib.MTCommand_TeXToggle')) | Out-Null
    if($doc.InlineShapes.Count -ne $Expected){throw "MathType object count mismatch: $($doc.InlineShapes.Count)/$Expected"}
    foreach($shape in $doc.InlineShapes){
        if($shape.OLEFormat.ProgID -ne 'Equation.DSMT4'){throw 'Unexpected equation object type'}
    }
    $doc.Save()
} finally {
    if($doc){$doc.Close(0)}
    if($word){$word.Quit()}
    if($doc){[Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null}
    if($word){[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
