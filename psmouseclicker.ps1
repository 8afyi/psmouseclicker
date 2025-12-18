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

while ($true) {
    [M]::mouse_event(2,0,0,0,0)
    [M]::mouse_event(4,0,0,0,0)

    $count++
    Write-Host ("Clicks: {0}" -f $count)

    Start-Sleep -Milliseconds 400
}
