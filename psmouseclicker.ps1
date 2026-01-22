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
$intro = @"
      _._
   .-'   `
 __|__
/     \
|()_()|
\{o o}/   Welcome to the PS Mouse Clicker
 =\o/=
  ^ ^

"@
Write-Host $intro -ForegroundColor Green

$delayInput = Read-Host "Enter delay in milliseconds between clicks (default 400)"

if ([int]::TryParse($delayInput, [ref]$parsedDelay) -and $parsedDelay -gt 0) {
    $delay = $parsedDelay
}

Write-Host "Using delay $delay ms" -ForegroundColor Black -BackgroundColor White
Write-Host "Press ESC to stop" -ForegroundColor Yellow

$spinner = @('|','/','-','\')
$spinnerIndex = 0
$rand = [Random]::new()
$exitRequested = $false

while (-not $exitRequested) {
    [M]::mouse_event(2,0,0,0,0)
    [M]::mouse_event(4,0,0,0,0)

    $count++
    $spinnerChar = $spinner[$spinnerIndex % $spinner.Length]
    $spinnerIndex++
    $maxJitter = [math]::Floor($delay * 0.1)
    $jitter = if ($maxJitter -gt 0) { $rand.Next(-$maxJitter, $maxJitter + 1) } else { 0 }
    $effectiveDelay = [math]::Max(1, $delay + $jitter)

    Write-Host ("`r[{0}] {1} Clicks: {2} | Delay: {3} ms" -f (Get-Date -Format "HH:mm:ss:ffff"), $spinnerChar, $count, $effectiveDelay) -NoNewline -ForegroundColor Cyan

    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Escape') {
            $exitRequested = $true
            break
        }
    }

    $slept = 0
    while ($slept -lt $effectiveDelay -and -not $exitRequested) {
        $chunk = [math]::Min(50, $effectiveDelay - $slept)
        Start-Sleep -Milliseconds $chunk
        $slept += $chunk

        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Escape') {
                $exitRequested = $true
                break
            }
        }
    }
}

if ($exitRequested) {
    # Clear the status line before printing exit messages.
    Write-Host ("`r" + (" " * 80)) -NoNewline
    Write-Host ("`r...  DONE!") -ForegroundColor Cyan
    Write-Host "ESC pressed.  Exiting."  -ForegroundColor Black -BackgroundColor White
}
