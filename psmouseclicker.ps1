param(
    [ValidateSet('Console', 'Gui')]
    [string]$Mode = 'Console',

    [Nullable[int]]$Delay,
    [switch]$DisableJitter,

    [ValidateRange(0, 86400)]
    [int]$StartDelaySec = 0,

    [Nullable[int]]$DurationSec,
    [Nullable[int]]$ClickLimit,

    [ValidateRange(0, 86400)]
    [int]$IdleTimeoutSec = 300,

    [ValidateRange(1, 200)]
    [double]$MaxCpsWarningCps = 15,

    [switch]$NoConfirm
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-NativeInterop {
    if (-not ([System.Management.Automation.PSTypeName]'NativeInput').Type) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct INPUT
{
    public uint type;
    public INPUTUNION U;
}

[StructLayout(LayoutKind.Explicit)]
public struct INPUTUNION
{
    [FieldOffset(0)]
    public MOUSEINPUT mi;
}

[StructLayout(LayoutKind.Sequential)]
public struct MOUSEINPUT
{
    public int dx;
    public int dy;
    public uint mouseData;
    public uint dwFlags;
    public uint time;
    public UIntPtr dwExtraInfo;
}

public static class NativeInput
{
    public const uint INPUT_MOUSE = 0;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
}
'@
    }
}

$script:ClickInputs = $null

function Initialize-ClickInputs {
    if ($null -ne $script:ClickInputs) {
        return
    }

    $downMouse = [MOUSEINPUT]::new()
    $downMouse.dwFlags = [NativeInput]::MOUSEEVENTF_LEFTDOWN
    $downUnion = [INPUTUNION]::new()
    $downUnion.mi = $downMouse

    $upMouse = [MOUSEINPUT]::new()
    $upMouse.dwFlags = [NativeInput]::MOUSEEVENTF_LEFTUP
    $upUnion = [INPUTUNION]::new()
    $upUnion.mi = $upMouse

    $downInput = [INPUT]::new()
    $downInput.type = [NativeInput]::INPUT_MOUSE
    $downInput.U = $downUnion

    $upInput = [INPUT]::new()
    $upInput.type = [NativeInput]::INPUT_MOUSE
    $upInput.U = $upUnion

    $script:ClickInputs = [INPUT[]]($downInput, $upInput)
}

function Invoke-LeftClick {
    $inputSize = [Runtime.InteropServices.Marshal]::SizeOf([INPUT])
    $sentCount = [NativeInput]::SendInput(2, $script:ClickInputs, $inputSize)
    if ($sentCount -ne 2) {
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "SendInput failed. Sent=$sentCount, LastError=$lastError"
    }
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

function Assert-OptionalPositiveInt {
    param(
        [string]$Name,
        [Nullable[int]]$Value
    )

    if ($null -eq $Value) {
        return
    }

    if ($Value -lt 1) {
        throw "$Name must be a positive integer when provided."
    }
}

function Format-OptionalLimit {
    param(
        [Nullable[int]]$Value,
        [string]$DisabledValue = 'None'
    )

    if ($null -eq $Value) {
        return $DisabledValue
    }

    return $Value.ToString()
}

function Format-IdleLimit {
    param([int]$Value)

    if ($Value -le 0) {
        return 'Disabled'
    }

    return "$Value s"
}

function Get-EstimatedCps {
    param([int]$DelayMs)
    return [math]::Round((1000.0 / [double][math]::Max(1, $DelayMs)), 2)
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

function Read-ConsoleKey {
    param([bool]$CanReadKeys)

    $result = [ordered]@{
        Read   = $false
        Escape = $false
    }

    if (-not $CanReadKeys) {
        return [pscustomobject]$result
    }

    try {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $result.Read = $true
            $result.Escape = ($key.Key -eq [ConsoleKey]::Escape)
        }
    }
    catch {
        $result.Read = $false
        $result.Escape = $false
    }

    return [pscustomobject]$result
}

function Confirm-ConsoleStart {
    param(
        [int]$DelayMs,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds,
        [double]$EstimatedCps,
        [double]$MaxCpsWarningThreshold,
        [bool]$RequireConfirm
    )

    Write-Host "Settings:" -ForegroundColor White
    Write-Host ("  Delay: {0} ms ({1} CPS estimated)" -f $DelayMs, $EstimatedCps) -ForegroundColor Gray
    Write-Host ("  Jitter: {0}" -f ($(if ($UseJitter) { 'Enabled (+/-10%)' } else { 'Disabled' }))) -ForegroundColor Gray
    Write-Host ("  Start delay: {0}s" -f $StartDelaySeconds) -ForegroundColor Gray
    Write-Host ("  Duration limit: {0}" -f (Format-OptionalLimit -Value $DurationLimitSeconds)) -ForegroundColor Gray
    Write-Host ("  Click limit: {0}" -f (Format-OptionalLimit -Value $ClickLimitValue)) -ForegroundColor Gray
    Write-Host ("  Idle timeout: {0}" -f (Format-IdleLimit -Value $IdleTimeoutSeconds)) -ForegroundColor Gray

    if ($EstimatedCps -gt $MaxCpsWarningThreshold) {
        Write-Warning ("Requested click rate is {0} CPS, above warning threshold {1} CPS." -f $EstimatedCps, $MaxCpsWarningThreshold)
        $fastAck = Read-Host "Type FAST to continue at this rate, or anything else to cancel"
        if ($fastAck -cne 'FAST') {
            return $false
        }
    }

    if (-not $RequireConfirm) {
        return $true
    }

    $confirmation = Read-Host "Type START to begin clicking"
    return ($confirmation -ceq 'START')
}

function Start-ConsoleClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds,
        [double]$MaxCpsWarningThreshold,
        [bool]$RequireConfirm
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
    $canReadKeys = Test-ConsoleKeySupport
    $estimatedCps = Get-EstimatedCps -DelayMs $delay

    Write-Host "Press ESC to stop (or Ctrl+C)." -ForegroundColor Yellow
    if (-not $canReadKeys) {
        Write-Host "ESC key detection is unavailable in this host." -ForegroundColor Yellow
    }

    $confirmed = Confirm-ConsoleStart `
        -DelayMs $delay `
        -UseJitter $UseJitter `
        -StartDelaySeconds $StartDelaySeconds `
        -DurationLimitSeconds $DurationLimitSeconds `
        -ClickLimitValue $ClickLimitValue `
        -IdleTimeoutSeconds $IdleTimeoutSeconds `
        -EstimatedCps $estimatedCps `
        -MaxCpsWarningThreshold $MaxCpsWarningThreshold `
        -RequireConfirm $RequireConfirm

    if (-not $confirmed) {
        Write-Host "Start canceled." -ForegroundColor Yellow
        return
    }

    $lastInteractionUtc = [DateTime]::UtcNow
    if ($StartDelaySeconds -gt 0) {
        for ($remaining = $StartDelaySeconds; $remaining -gt 0; $remaining--) {
            Write-Host ("`rStarting in {0}s. Press ESC to cancel..." -f $remaining) -NoNewline -ForegroundColor DarkYellow

            $waitedMs = 0
            while ($waitedMs -lt 1000) {
                Start-Sleep -Milliseconds 100
                $waitedMs += 100

                $keyInfo = Read-ConsoleKey -CanReadKeys $canReadKeys
                if ($keyInfo.Read) {
                    $lastInteractionUtc = [DateTime]::UtcNow
                }

                if ($keyInfo.Escape) {
                    Write-Host ""
                    Write-Host "Canceled before start." -ForegroundColor Yellow
                    return
                }
            }
        }

        Write-Host ("`r" + (" " * 90)) -NoNewline
        Write-Host ("`rStarting...") -ForegroundColor Green
    }

    $runStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $stopReason = $null

    :MainLoop while ($true) {
        if ($null -ne $DurationLimitSeconds -and $runStopwatch.Elapsed.TotalSeconds -ge $DurationLimitSeconds) {
            $stopReason = "Duration limit reached ($DurationLimitSeconds s)"
            break
        }

        if ($null -ne $ClickLimitValue -and $count -ge $ClickLimitValue) {
            $stopReason = "Click limit reached ($ClickLimitValue)"
            break
        }

        if ($IdleTimeoutSeconds -gt 0 -and ([DateTime]::UtcNow - $lastInteractionUtc).TotalSeconds -ge $IdleTimeoutSeconds) {
            $stopReason = "Idle timeout reached ($IdleTimeoutSeconds s)"
            break
        }

        Invoke-LeftClick
        $count++

        $effectiveDelay = Get-EffectiveDelay -BaseDelay $delay -UseJitter $UseJitter -Random $rand
        $spinnerChar = $spinner[$spinnerIndex % $spinner.Length]
        $spinnerIndex++

        $statusParts = @(
            ("Clicks: {0}" -f $count),
            ("Delay: {0} ms" -f $effectiveDelay)
        )

        if ($null -ne $ClickLimitValue) {
            $statusParts += ("Remaining clicks: {0}" -f [math]::Max(0, ($ClickLimitValue - $count)))
        }

        if ($null -ne $DurationLimitSeconds) {
            $remainingSeconds = [math]::Max(0, [int][math]::Ceiling($DurationLimitSeconds - $runStopwatch.Elapsed.TotalSeconds))
            $statusParts += ("Time left: {0}s" -f $remainingSeconds)
        }

        if ($IdleTimeoutSeconds -gt 0) {
            $idleLeft = [math]::Max(0, [int][math]::Ceiling($IdleTimeoutSeconds - ([DateTime]::UtcNow - $lastInteractionUtc).TotalSeconds))
            $statusParts += ("Idle left: {0}s" -f $idleLeft)
        }

        Write-Host ("`r[{0}] {1} {2}" -f (Get-Date -Format "HH:mm:ss.fff"), $spinnerChar, ($statusParts -join " | ")) -NoNewline -ForegroundColor Cyan

        $slept = 0
        while ($slept -lt $effectiveDelay) {
            $chunk = [math]::Min(50, ($effectiveDelay - $slept))
            Start-Sleep -Milliseconds $chunk
            $slept += $chunk

            $keyInfo = Read-ConsoleKey -CanReadKeys $canReadKeys
            if ($keyInfo.Read) {
                $lastInteractionUtc = [DateTime]::UtcNow
            }

            if ($keyInfo.Escape) {
                $stopReason = "ESC pressed"
                break MainLoop
            }

            if ($null -ne $DurationLimitSeconds -and $runStopwatch.Elapsed.TotalSeconds -ge $DurationLimitSeconds) {
                $stopReason = "Duration limit reached ($DurationLimitSeconds s)"
                break MainLoop
            }

            if ($IdleTimeoutSeconds -gt 0 -and ([DateTime]::UtcNow - $lastInteractionUtc).TotalSeconds -ge $IdleTimeoutSeconds) {
                $stopReason = "Idle timeout reached ($IdleTimeoutSeconds s)"
                break MainLoop
            }
        }
    }

    $lineWidth = 80
    try {
        $lineWidth = [math]::Max(20, [Console]::WindowWidth - 1)
    }
    catch {
        $lineWidth = 80
    }

    Write-Host ("`r" + (" " * $lineWidth)) -NoNewline
    Write-Host ("`r... DONE!") -ForegroundColor Cyan
    Write-Host ("Reason: {0}" -f $stopReason) -ForegroundColor Black -BackgroundColor White
    Write-Host ("Total clicks: {0}" -f $count) -ForegroundColor Gray
}

function Confirm-GuiStart {
    param(
        $Owner,
        [int]$DelayMs,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds,
        [double]$MaxCpsWarningThreshold,
        [bool]$RequireConfirm
    )

    $estimatedCps = Get-EstimatedCps -DelayMs $DelayMs
    if ($estimatedCps -gt $MaxCpsWarningThreshold) {
        $warningText = @"
The requested click rate is $estimatedCps CPS.
This exceeds the warning threshold of $MaxCpsWarningThreshold CPS.

Continue anyway?
"@
        $rateResult = [System.Windows.Forms.MessageBox]::Show(
            $Owner,
            $warningText,
            "High Click Rate Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($rateResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return $false
        }
    }

    if (-not $RequireConfirm) {
        return $true
    }

    $summary = @"
Start clicking with these settings?

Delay: $DelayMs ms
Jitter: $(if ($UseJitter) { 'Enabled (+/-10%)' } else { 'Disabled' })
Start delay: ${StartDelaySeconds}s
Duration limit: $(Format-OptionalLimit -Value $DurationLimitSeconds)
Click limit: $(Format-OptionalLimit -Value $ClickLimitValue)
Idle timeout: $(Format-IdleLimit -Value $IdleTimeoutSeconds)
"@

    $result = [System.Windows.Forms.MessageBox]::Show(
        $Owner,
        $summary,
        "Confirm Start",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Start-GuiClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds,
        [double]$MaxCpsWarningThreshold,
        [bool]$RequireConfirm
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
        throw "GUI mode requires an STA thread. Run powershell.exe, or start pwsh with -STA, or use -Mode Console."
    }

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $state = @{
        ClickCount          = 0
        Phase               = 'Stopped'
        Random              = [Random]::new()
        RunStartedUtc       = $null
        PendingStartUtc     = $null
        LastInteractionUtc  = [DateTime]::UtcNow
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PS Mouse Clicker"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ClientSize = New-Object System.Drawing.Size(460, 320)
    $form.KeyPreview = $true

    $uiFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelWidth = 160
    $valueX = 190

    $delayLabel = New-Object System.Windows.Forms.Label
    $delayLabel.Text = "Delay between clicks:"
    $delayLabel.Location = New-Object System.Drawing.Point(20, 20)
    $delayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $delayLabel.Font = $uiFont

    $delayInput = New-Object System.Windows.Forms.NumericUpDown
    $delayInput.Location = New-Object System.Drawing.Point($valueX, 18)
    $delayInput.Size = New-Object System.Drawing.Size(100, 24)
    $delayInput.Minimum = 1
    $delayInput.Maximum = 60000
    $delayInput.Value = [math]::Min(60000, [math]::Max(1, $InitialDelay))
    $delayInput.Font = $uiFont

    $delayUnits = New-Object System.Windows.Forms.Label
    $delayUnits.Text = "ms"
    $delayUnits.Location = New-Object System.Drawing.Point(300, 20)
    $delayUnits.Size = New-Object System.Drawing.Size(60, 24)
    $delayUnits.Font = $uiFont

    $jitterCheckBox = New-Object System.Windows.Forms.CheckBox
    $jitterCheckBox.Text = "Enable +/-10% jitter"
    $jitterCheckBox.Location = New-Object System.Drawing.Point($valueX, 50)
    $jitterCheckBox.Size = New-Object System.Drawing.Size(220, 24)
    $jitterCheckBox.Checked = $UseJitter
    $jitterCheckBox.Font = $uiFont

    $limitsLabel = New-Object System.Windows.Forms.Label
    $limitsLabel.Text = "Run limits:"
    $limitsLabel.Location = New-Object System.Drawing.Point(20, 84)
    $limitsLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $limitsLabel.Font = $uiFont

    $limitsValue = New-Object System.Windows.Forms.Label
    $limitsValue.Text = @"
Start delay: ${StartDelaySeconds}s | Duration: $(Format-OptionalLimit -Value $DurationLimitSeconds)
Click limit: $(Format-OptionalLimit -Value $ClickLimitValue) | Idle timeout: $(Format-IdleLimit -Value $IdleTimeoutSeconds)
"@
    $limitsValue.Location = New-Object System.Drawing.Point($valueX, 82)
    $limitsValue.Size = New-Object System.Drawing.Size(250, 46)
    $limitsValue.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Status:"
    $statusLabel.Location = New-Object System.Drawing.Point(20, 136)
    $statusLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $statusLabel.Font = $uiFont

    $statusValue = New-Object System.Windows.Forms.Label
    $statusValue.Text = "Stopped"
    $statusValue.Location = New-Object System.Drawing.Point($valueX, 136)
    $statusValue.Size = New-Object System.Drawing.Size(240, 24)
    $statusValue.Font = $uiFont

    $clicksLabel = New-Object System.Windows.Forms.Label
    $clicksLabel.Text = "Clicks:"
    $clicksLabel.Location = New-Object System.Drawing.Point(20, 166)
    $clicksLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $clicksLabel.Font = $uiFont

    $clicksValue = New-Object System.Windows.Forms.Label
    $clicksValue.Text = "0"
    $clicksValue.Location = New-Object System.Drawing.Point($valueX, 166)
    $clicksValue.Size = New-Object System.Drawing.Size(240, 24)
    $clicksValue.Font = $uiFont

    $currentDelayLabel = New-Object System.Windows.Forms.Label
    $currentDelayLabel.Text = "Current delay:"
    $currentDelayLabel.Location = New-Object System.Drawing.Point(20, 196)
    $currentDelayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $currentDelayLabel.Font = $uiFont

    $currentDelayValue = New-Object System.Windows.Forms.Label
    $currentDelayValue.Text = "{0} ms" -f [int]$delayInput.Value
    $currentDelayValue.Location = New-Object System.Drawing.Point($valueX, 196)
    $currentDelayValue.Size = New-Object System.Drawing.Size(240, 24)
    $currentDelayValue.Font = $uiFont

    $lastClickLabel = New-Object System.Windows.Forms.Label
    $lastClickLabel.Text = "Last click:"
    $lastClickLabel.Location = New-Object System.Drawing.Point(20, 226)
    $lastClickLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $lastClickLabel.Font = $uiFont

    $lastClickValue = New-Object System.Windows.Forms.Label
    $lastClickValue.Text = "-"
    $lastClickValue.Location = New-Object System.Drawing.Point($valueX, 226)
    $lastClickValue.Size = New-Object System.Drawing.Size(240, 24)
    $lastClickValue.Font = $uiFont

    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "Start"
    $startButton.Location = New-Object System.Drawing.Point(20, 272)
    $startButton.Size = New-Object System.Drawing.Size(130, 32)
    $startButton.Font = $uiFont

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Text = "Stop"
    $stopButton.Location = New-Object System.Drawing.Point(165, 272)
    $stopButton.Size = New-Object System.Drawing.Size(130, 32)
    $stopButton.Enabled = $false
    $stopButton.Font = $uiFont

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset Count"
    $resetButton.Location = New-Object System.Drawing.Point(310, 272)
    $resetButton.Size = New-Object System.Drawing.Size(130, 32)
    $resetButton.Font = $uiFont

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [int]$delayInput.Value

    $touchInteraction = {
        $state.LastInteractionUtc = [DateTime]::UtcNow
    }

    $stopClicker = {
        param([string]$Reason = "Stopped")

        $timer.Stop()
        $state.Phase = 'Stopped'
        $state.RunStartedUtc = $null
        $state.PendingStartUtc = $null

        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        $delayInput.Enabled = $true
        $jitterCheckBox.Enabled = $true
        $statusValue.Text = $Reason
    }

    $startButton.Add_Click({
        & $touchInteraction

        if ($state.Phase -ne 'Stopped') {
            return
        }

        $baseDelay = [int]$delayInput.Value
        $approved = Confirm-GuiStart `
            -Owner $form `
            -DelayMs $baseDelay `
            -UseJitter $jitterCheckBox.Checked `
            -StartDelaySeconds $StartDelaySeconds `
            -DurationLimitSeconds $DurationLimitSeconds `
            -ClickLimitValue $ClickLimitValue `
            -IdleTimeoutSeconds $IdleTimeoutSeconds `
            -MaxCpsWarningThreshold $MaxCpsWarningThreshold `
            -RequireConfirm $RequireConfirm

        if (-not $approved) {
            $statusValue.Text = "Start canceled"
            return
        }

        $startButton.Enabled = $false
        $stopButton.Enabled = $true
        $delayInput.Enabled = $false
        $jitterCheckBox.Enabled = $false
        $state.LastInteractionUtc = [DateTime]::UtcNow

        if ($StartDelaySeconds -gt 0) {
            $state.Phase = 'Countdown'
            $state.PendingStartUtc = [DateTime]::UtcNow.AddSeconds($StartDelaySeconds)
            $statusValue.Text = "Starting in ${StartDelaySeconds}s"
            $timer.Interval = 200
        }
        else {
            $state.Phase = 'Running'
            $state.RunStartedUtc = [DateTime]::UtcNow
            $statusValue.Text = "Running"
            $timer.Interval = [int]$delayInput.Value
        }

        $timer.Start()
    })

    $stopButton.Add_Click({
        & $touchInteraction
        & $stopClicker "Stopped by user"
    })

    $resetButton.Add_Click({
        & $touchInteraction
        $state.ClickCount = 0
        $clicksValue.Text = "0"
        $lastClickValue.Text = "-"
    })

    $delayInput.Add_ValueChanged({
        & $touchInteraction
        if ($state.Phase -eq 'Stopped') {
            $currentDelayValue.Text = "{0} ms" -f [int]$delayInput.Value
        }
    })

    $timer.Add_Tick({
        $now = [DateTime]::UtcNow

        if ($IdleTimeoutSeconds -gt 0 -and ($now - $state.LastInteractionUtc).TotalSeconds -ge $IdleTimeoutSeconds) {
            & $stopClicker ("Idle timeout reached ({0}s)" -f $IdleTimeoutSeconds)
            return
        }

        if ($state.Phase -eq 'Countdown') {
            $remainingMs = [int][math]::Ceiling(($state.PendingStartUtc - $now).TotalMilliseconds)
            if ($remainingMs -le 0) {
                $state.Phase = 'Running'
                $state.RunStartedUtc = [DateTime]::UtcNow
                $statusValue.Text = "Running"
                $timer.Interval = [int]$delayInput.Value
                return
            }

            $remainingSec = [int][math]::Ceiling($remainingMs / 1000.0)
            $statusValue.Text = "Starting in ${remainingSec}s"
            return
        }

        if ($state.Phase -ne 'Running') {
            return
        }

        if ($null -ne $DurationLimitSeconds -and ($now - $state.RunStartedUtc).TotalSeconds -ge $DurationLimitSeconds) {
            & $stopClicker ("Duration limit reached ({0}s)" -f $DurationLimitSeconds)
            return
        }

        if ($null -ne $ClickLimitValue -and $state.ClickCount -ge $ClickLimitValue) {
            & $stopClicker ("Click limit reached ({0})" -f $ClickLimitValue)
            return
        }

        Invoke-LeftClick
        $state.ClickCount++
        $clicksValue.Text = $state.ClickCount.ToString()

        $baseDelay = [int]$delayInput.Value
        $effectiveDelay = Get-EffectiveDelay -BaseDelay $baseDelay -UseJitter $jitterCheckBox.Checked -Random $state.Random
        $currentDelayValue.Text = "{0} ms" -f $effectiveDelay
        $lastClickValue.Text = Get-Date -Format "HH:mm:ss.fff"
        $timer.Interval = $effectiveDelay

        if ($null -ne $ClickLimitValue -and $state.ClickCount -ge $ClickLimitValue) {
            & $stopClicker ("Click limit reached ({0})" -f $ClickLimitValue)
        }
    })

    $form.Add_KeyDown({
        param($sender, $eventArgs)

        & $touchInteraction
        if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            & $stopClicker "Stopped (ESC)"
            $eventArgs.SuppressKeyPress = $true
        }
    })

    foreach ($control in @($form, $delayInput, $jitterCheckBox, $startButton, $stopButton, $resetButton)) {
        $control.Add_MouseMove({
            & $touchInteraction
        })
    }

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
            $limitsLabel,
            $limitsValue,
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

Assert-OptionalPositiveInt -Name 'DurationSec' -Value $DurationSec
Assert-OptionalPositiveInt -Name 'ClickLimit' -Value $ClickLimit

Ensure-NativeInterop
Initialize-ClickInputs

$useJitter = -not $DisableJitter
$initialDelay = Resolve-InitialDelay -DelayValue $Delay -PromptInConsole ($Mode -eq 'Console')
$requireConfirm = -not $NoConfirm

switch ($Mode) {
    'Console' {
        Start-ConsoleClicker `
            -InitialDelay $initialDelay `
            -UseJitter $useJitter `
            -StartDelaySeconds $StartDelaySec `
            -DurationLimitSeconds $DurationSec `
            -ClickLimitValue $ClickLimit `
            -IdleTimeoutSeconds $IdleTimeoutSec `
            -MaxCpsWarningThreshold $MaxCpsWarningCps `
            -RequireConfirm $requireConfirm
    }
    'Gui' {
        Start-GuiClicker `
            -InitialDelay $initialDelay `
            -UseJitter $useJitter `
            -StartDelaySeconds $StartDelaySec `
            -DurationLimitSeconds $DurationSec `
            -ClickLimitValue $ClickLimit `
            -IdleTimeoutSeconds $IdleTimeoutSec `
            -MaxCpsWarningThreshold $MaxCpsWarningCps `
            -RequireConfirm $requireConfirm
    }
}
