param(
    [ValidateSet('Console', 'Gui')]
    [string]$Mode = 'Console',
    [Nullable[int]]$Delay,
    [switch]$DisableJitter
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-MouseInterop {
    if (-not ([System.Management.Automation.PSTypeName]'MouseInterop').Type) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class MouseInterop
{
    public const uint LeftDown = 0x0002;
    public const uint LeftUp = 0x0004;

    [DllImport("user32.dll", SetLastError=true)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
}
'@
    }
}

function Invoke-LeftClick {
    [MouseInterop]::mouse_event([MouseInterop]::LeftDown, 0, 0, 0, [UIntPtr]::Zero)
    [MouseInterop]::mouse_event([MouseInterop]::LeftUp, 0, 0, 0, [UIntPtr]::Zero)
}

function Get-EffectiveDelay {
    param(
        [int]$BaseDelay,
        [bool]$UseJitter,
        [Random]$Random
    )

    if (-not $UseJitter) {
        return [math]::Max(1, $BaseDelay)
    }

    $maxJitter = [math]::Floor($BaseDelay * 0.1)
    if ($maxJitter -le 0) {
        return [math]::Max(1, $BaseDelay)
    }

    $jitter = $Random.Next(-$maxJitter, $maxJitter + 1)
    return [math]::Max(1, $BaseDelay + $jitter)
}

function Resolve-InitialDelay {
    param(
        [Nullable[int]]$DelayValue,
        [bool]$PromptInConsole
    )

    $defaultDelay = 400

    if ($null -ne $DelayValue) {
        if ($DelayValue -lt 1) {
            throw "Delay must be a positive integer."
        }

        return [int]$DelayValue
    }

    if (-not $PromptInConsole) {
        return $defaultDelay
    }

    $parsedDelay = 0
    $delayInput = Read-Host "Enter delay in milliseconds between clicks (default $defaultDelay)"
    if ([int]::TryParse($delayInput, [ref]$parsedDelay) -and $parsedDelay -gt 0) {
        return $parsedDelay
    }

    return $defaultDelay
}

function Test-ConsoleKeySupport {
    try {
        [void][Console]::KeyAvailable
        return $true
    }
    catch {
        return $false
    }
}

function Test-EscapePressed {
    param([bool]$CanReadKeys)

    if (-not $CanReadKeys) {
        return $false
    }

    try {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            return ($key.Key -eq [ConsoleKey]::Escape)
        }

        return $false
    }
    catch {
        return $false
    }
}

function Start-ConsoleClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter
    )

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

    $delay = $InitialDelay
    $count = 0
    $spinner = @('|', '/', '-', '\')
    $spinnerIndex = 0
    $rand = [Random]::new()
    $exitRequested = $false
    $canReadKeys = Test-ConsoleKeySupport

    Write-Host "Using delay $delay ms" -ForegroundColor Black -BackgroundColor White
    Write-Host ("Jitter: {0}" -f ($(if ($UseJitter) { 'Enabled (+/-10%)' } else { 'Disabled' }))) -ForegroundColor DarkGray

    if ($canReadKeys) {
        Write-Host "Press ESC to stop" -ForegroundColor Yellow
    }
    else {
        Write-Host "ESC detection is unavailable in this host. Press Ctrl+C to stop." -ForegroundColor Yellow
    }

    while (-not $exitRequested) {
        Invoke-LeftClick
        $count++

        $effectiveDelay = Get-EffectiveDelay -BaseDelay $delay -UseJitter $UseJitter -Random $rand
        $spinnerChar = $spinner[$spinnerIndex % $spinner.Length]
        $spinnerIndex++

        Write-Host ("`r[{0}] {1} Clicks: {2} | Delay: {3} ms" -f (Get-Date -Format "HH:mm:ss.fff"), $spinnerChar, $count, $effectiveDelay) -NoNewline -ForegroundColor Cyan

        if (Test-EscapePressed -CanReadKeys $canReadKeys) {
            $exitRequested = $true
            break
        }

        $slept = 0
        while ($slept -lt $effectiveDelay -and -not $exitRequested) {
            $chunk = [math]::Min(50, $effectiveDelay - $slept)
            Start-Sleep -Milliseconds $chunk
            $slept += $chunk

            if (Test-EscapePressed -CanReadKeys $canReadKeys) {
                $exitRequested = $true
                break
            }
        }
    }

    if ($exitRequested) {
        $lineWidth = 80
        try {
            $lineWidth = [math]::Max(20, [Console]::WindowWidth - 1)
        }
        catch {
            $lineWidth = 80
        }

        Write-Host ("`r" + (" " * $lineWidth)) -NoNewline
        Write-Host ("`r...  DONE!") -ForegroundColor Cyan
        Write-Host "ESC pressed. Exiting." -ForegroundColor Black -BackgroundColor White
    }
}

function Start-GuiClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
        throw "GUI mode requires an STA thread. Run powershell.exe, or start pwsh with -STA, or use -Mode Console."
    }

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $state = @{
        ClickCount = 0
        IsRunning  = $false
        Random     = [Random]::new()
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PS Mouse Clicker"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ClientSize = New-Object System.Drawing.Size(420, 260)
    $form.KeyPreview = $true

    $uiFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelWidth = 140
    $valueX = 170

    $delayLabel = New-Object System.Windows.Forms.Label
    $delayLabel.Text = "Delay between clicks:"
    $delayLabel.Location = New-Object System.Drawing.Point(20, 24)
    $delayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $delayLabel.Font = $uiFont

    $delayInput = New-Object System.Windows.Forms.NumericUpDown
    $delayInput.Location = New-Object System.Drawing.Point($valueX, 22)
    $delayInput.Size = New-Object System.Drawing.Size(100, 24)
    $delayInput.Minimum = 1
    $delayInput.Maximum = 60000
    $delayInput.Value = [math]::Min(60000, [math]::Max(1, $InitialDelay))
    $delayInput.Font = $uiFont

    $delayUnits = New-Object System.Windows.Forms.Label
    $delayUnits.Text = "ms"
    $delayUnits.Location = New-Object System.Drawing.Point(280, 24)
    $delayUnits.Size = New-Object System.Drawing.Size(60, 24)
    $delayUnits.Font = $uiFont

    $jitterCheckBox = New-Object System.Windows.Forms.CheckBox
    $jitterCheckBox.Text = "Enable +/-10% jitter"
    $jitterCheckBox.Location = New-Object System.Drawing.Point($valueX, 56)
    $jitterCheckBox.Size = New-Object System.Drawing.Size(200, 24)
    $jitterCheckBox.Checked = $UseJitter
    $jitterCheckBox.Font = $uiFont

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Status:"
    $statusLabel.Location = New-Object System.Drawing.Point(20, 92)
    $statusLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $statusLabel.Font = $uiFont

    $statusValue = New-Object System.Windows.Forms.Label
    $statusValue.Text = "Stopped"
    $statusValue.Location = New-Object System.Drawing.Point($valueX, 92)
    $statusValue.Size = New-Object System.Drawing.Size(220, 24)
    $statusValue.Font = $uiFont

    $clicksLabel = New-Object System.Windows.Forms.Label
    $clicksLabel.Text = "Clicks:"
    $clicksLabel.Location = New-Object System.Drawing.Point(20, 122)
    $clicksLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $clicksLabel.Font = $uiFont

    $clicksValue = New-Object System.Windows.Forms.Label
    $clicksValue.Text = "0"
    $clicksValue.Location = New-Object System.Drawing.Point($valueX, 122)
    $clicksValue.Size = New-Object System.Drawing.Size(220, 24)
    $clicksValue.Font = $uiFont

    $currentDelayLabel = New-Object System.Windows.Forms.Label
    $currentDelayLabel.Text = "Current delay:"
    $currentDelayLabel.Location = New-Object System.Drawing.Point(20, 152)
    $currentDelayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $currentDelayLabel.Font = $uiFont

    $currentDelayValue = New-Object System.Windows.Forms.Label
    $currentDelayValue.Text = "{0} ms" -f [int]$delayInput.Value
    $currentDelayValue.Location = New-Object System.Drawing.Point($valueX, 152)
    $currentDelayValue.Size = New-Object System.Drawing.Size(220, 24)
    $currentDelayValue.Font = $uiFont

    $lastClickLabel = New-Object System.Windows.Forms.Label
    $lastClickLabel.Text = "Last click:"
    $lastClickLabel.Location = New-Object System.Drawing.Point(20, 182)
    $lastClickLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $lastClickLabel.Font = $uiFont

    $lastClickValue = New-Object System.Windows.Forms.Label
    $lastClickValue.Text = "-"
    $lastClickValue.Location = New-Object System.Drawing.Point($valueX, 182)
    $lastClickValue.Size = New-Object System.Drawing.Size(220, 24)
    $lastClickValue.Font = $uiFont

    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "Start"
    $startButton.Location = New-Object System.Drawing.Point(20, 218)
    $startButton.Size = New-Object System.Drawing.Size(120, 30)
    $startButton.Font = $uiFont

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Text = "Stop"
    $stopButton.Location = New-Object System.Drawing.Point(150, 218)
    $stopButton.Size = New-Object System.Drawing.Size(120, 30)
    $stopButton.Enabled = $false
    $stopButton.Font = $uiFont

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset Count"
    $resetButton.Location = New-Object System.Drawing.Point(280, 218)
    $resetButton.Size = New-Object System.Drawing.Size(120, 30)
    $resetButton.Font = $uiFont

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [int]$delayInput.Value

    $stopClicker = {
        if (-not $state.IsRunning) {
            return
        }

        $state.IsRunning = $false
        $timer.Stop()
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        $delayInput.Enabled = $true
        $jitterCheckBox.Enabled = $true
        $statusValue.Text = "Stopped"
    }

    $startButton.Add_Click({
        if ($state.IsRunning) {
            return
        }

        $state.IsRunning = $true
        $startButton.Enabled = $false
        $stopButton.Enabled = $true
        $delayInput.Enabled = $false
        $jitterCheckBox.Enabled = $false
        $statusValue.Text = "Running"
        $timer.Interval = [int]$delayInput.Value
        $timer.Start()
    })

    $stopButton.Add_Click({
        & $stopClicker
    })

    $resetButton.Add_Click({
        $state.ClickCount = 0
        $clicksValue.Text = "0"
        $lastClickValue.Text = "-"
    })

    $delayInput.Add_ValueChanged({
        if (-not $state.IsRunning) {
            $currentDelayValue.Text = "{0} ms" -f [int]$delayInput.Value
        }
    })

    $timer.Add_Tick({
        Invoke-LeftClick
        $state.ClickCount++
        $clicksValue.Text = $state.ClickCount.ToString()

        $baseDelay = [int]$delayInput.Value
        $effectiveDelay = Get-EffectiveDelay -BaseDelay $baseDelay -UseJitter $jitterCheckBox.Checked -Random $state.Random
        $currentDelayValue.Text = "{0} ms" -f $effectiveDelay
        $lastClickValue.Text = Get-Date -Format "HH:mm:ss.fff"
        $timer.Interval = $effectiveDelay
    })

    $form.Add_KeyDown({
        param($sender, $eventArgs)

        if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            & $stopClicker
            if (-not $state.IsRunning) {
                $statusValue.Text = "Stopped (ESC)"
            }
            $eventArgs.SuppressKeyPress = $true
        }
    })

    $form.Add_FormClosing({
        $timer.Stop()
        $timer.Dispose()
        $uiFont.Dispose()
    })

    $form.Controls.AddRange(@(
            $delayLabel,
            $delayInput,
            $delayUnits,
            $jitterCheckBox,
            $statusLabel,
            $statusValue,
            $clicksLabel,
            $clicksValue,
            $currentDelayLabel,
            $currentDelayValue,
            $lastClickLabel,
            $lastClickValue,
            $startButton,
            $stopButton,
            $resetButton
        ))

    [void]$form.ShowDialog()
}

Ensure-MouseInterop

$useJitter = -not $DisableJitter
$initialDelay = Resolve-InitialDelay -DelayValue $Delay -PromptInConsole ($Mode -eq 'Console')

switch ($Mode) {
    'Console' {
        Start-ConsoleClicker -InitialDelay $initialDelay -UseJitter $useJitter
    }
    'Gui' {
        Start-GuiClicker -InitialDelay $initialDelay -UseJitter $useJitter
    }
}
