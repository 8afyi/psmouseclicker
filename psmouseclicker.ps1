if (-not ([System.Management.Automation.PSTypeName]'M').Type) {
    # Define the interop type once so reruns don't raise TYPE_ALREADY_EXISTS.
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class M
{
    [DllImport("user32.dll")]
    public static extern void mouse_event(int a, int b, int c, int d, int e);
}
'@
}

$count = 0
$delay = 400
$parsedDelay = 0
$delayInput = Read-Host "Enter delay in milliseconds between clicks (default 400)"

if ([int]::TryParse($delayInput, [ref]$parsedDelay) -and $parsedDelay -gt 0) {
    $delay = $parsedDelay
}

Write-Host "Using delay $delay ms"
Write-Host "Press ESC to stop" -ForegroundColor Yellow

$spinner = @('|','/','-','\')
$spinnerIndex = 0

while ($true) {
    [M]::mouse_event(2,0,0,0,0)
    [M]::mouse_event(4,0,0,0,0)

    $count++
    $spinnerChar = $spinner[$spinnerIndex % $spinner.Length]
    $spinnerIndex++
    Write-Host ("`r[{0}] {1} Clicks: {2}" -f (Get-Date -Format "HH:mm:ss:ffff"), $spinnerChar, $count) -NoNewline -ForegroundColor Cyan

    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Escape') {
            Write-Host "...  DONE!"
            Write-Host "ESC pressed.  Exiting."
            break
        }
    }

    Start-Sleep -Milliseconds $delay
}
