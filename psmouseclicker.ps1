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
    [int]$IdleTimeoutSec = 0,

    [string]$LifetimeClicksFile = 'lifetime-clicks.txt'
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
$script:LifetimeClicks = 0L
$script:LifetimeClicksDirty = $false
$script:LifetimeClicksFilePath = $null

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
    $inputSize = [Runtime.InteropServices.Marshal]::SizeOf([INPUT]::new())
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

function Resolve-LifetimeClicksFilePath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        throw "LifetimeClicksFile cannot be empty."
    }

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    $basePath = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        (Get-Location).Path
    }
    else {
        $PSScriptRoot
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $basePath -ChildPath $PathValue))
}

function Save-LifetimeClicks {
    param(
        [string]$FilePath,
        [long]$Value
    )

    $directory = Split-Path -Path $FilePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }

    Set-Content -LiteralPath $FilePath -Value $Value -Encoding Ascii
}

function Initialize-LifetimeClicks {
    param([string]$FilePathValue)

    $resolvedPath = Resolve-LifetimeClicksFilePath -PathValue $FilePathValue
    $script:LifetimeClicksFilePath = $resolvedPath

    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        Save-LifetimeClicks -FilePath $resolvedPath -Value 0
        $script:LifetimeClicks = 0
        $script:LifetimeClicksDirty = $false
        return
    }

    $rawValue = (Get-Content -LiteralPath $resolvedPath -Raw -ErrorAction Stop).Trim()
    if ([string]::IsNullOrWhiteSpace($rawValue)) {
        $script:LifetimeClicks = 0
        Save-LifetimeClicks -FilePath $resolvedPath -Value $script:LifetimeClicks
        $script:LifetimeClicksDirty = $false
        return
    }

    $parsedValue = 0L
    if (-not [long]::TryParse($rawValue, [ref]$parsedValue) -or $parsedValue -lt 0) {
        throw "Lifetime click file '$resolvedPath' is invalid. Expected a non-negative integer."
    }

    $script:LifetimeClicks = $parsedValue
    $script:LifetimeClicksDirty = $false
}

function Add-LifetimeClicks {
    param(
        [ValidateRange(1, 2147483647)]
        [int]$Count = 1
    )

    $script:LifetimeClicks += [long]$Count
    $script:LifetimeClicksDirty = $true
}

function Flush-LifetimeClicks {
    if (-not $script:LifetimeClicksDirty) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:LifetimeClicksFilePath)) {
        return
    }

    try {
        Save-LifetimeClicks -FilePath $script:LifetimeClicksFilePath -Value $script:LifetimeClicks
        $script:LifetimeClicksDirty = $false
    }
    catch {
        Write-Warning ("Failed to persist lifetime clicks to '{0}': {1}" -f $script:LifetimeClicksFilePath, $_.Exception.Message)
    }
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

function Start-ConsoleClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds
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

    Write-Host "Press ESC to stop (or Ctrl+C)." -ForegroundColor Yellow
    if (-not $canReadKeys) {
        Write-Host "ESC key detection is unavailable in this host." -ForegroundColor Yellow
    }

    Write-Host "Settings:" -ForegroundColor White
    Write-Host ("  Delay: {0} ms" -f $delay) -ForegroundColor Gray
    Write-Host ("  Jitter: {0}" -f ($(if ($UseJitter) { 'Enabled (+/-10%)' } else { 'Disabled' }))) -ForegroundColor Gray
    Write-Host ("  Start delay: {0}s" -f $StartDelaySeconds) -ForegroundColor Gray
    Write-Host ("  Duration limit: {0}" -f (Format-OptionalLimit -Value $DurationLimitSeconds)) -ForegroundColor Gray
    Write-Host ("  Click limit: {0}" -f (Format-OptionalLimit -Value $ClickLimitValue)) -ForegroundColor Gray
    Write-Host ("  Idle timeout: {0}" -f (Format-IdleLimit -Value $IdleTimeoutSeconds)) -ForegroundColor Gray
    Write-Host ("  Lifetime clicks (before run): {0}" -f $script:LifetimeClicks) -ForegroundColor Gray
    Write-Host ("  Lifetime file: {0}" -f $script:LifetimeClicksFilePath) -ForegroundColor DarkGray

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
    $flushEveryClicks = 50
    $clicksSinceFlush = 0

    try {
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
            Add-LifetimeClicks -Count 1
            $clicksSinceFlush++
            if ($clicksSinceFlush -ge $flushEveryClicks) {
                Flush-LifetimeClicks
                $clicksSinceFlush = 0
            }

            $effectiveDelay = Get-EffectiveDelay -BaseDelay $delay -UseJitter $UseJitter -Random $rand
            $spinnerChar = $spinner[$spinnerIndex % $spinner.Length]
            $spinnerIndex++

            $statusParts = @(
                ("Clicks: {0}" -f $count),
                ("Lifetime: {0}" -f $script:LifetimeClicks),
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
    }
    finally {
        Flush-LifetimeClicks
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
    Write-Host ("Lifetime clicks (all-time): {0}" -f $script:LifetimeClicks) -ForegroundColor Gray
}

function Start-GuiClicker {
    param(
        [int]$InitialDelay,
        [bool]$UseJitter,
        [int]$StartDelaySeconds,
        [Nullable[int]]$DurationLimitSeconds,
        [Nullable[int]]$ClickLimitValue,
        [int]$IdleTimeoutSeconds
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
        ClicksSinceFlush    = 0
        StartDelaySeconds   = $StartDelaySeconds
        DurationLimitSeconds = $DurationLimitSeconds
        ClickLimitValue     = $ClickLimitValue
        IdleTimeoutSeconds  = $IdleTimeoutSeconds
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PS Mouse Clicker"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.ClientSize = New-Object System.Drawing.Size(620, 440)
    $form.KeyPreview = $true

    $uiFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelWidth = 160
    $valueX = 190
    $flushEveryClicks = 50

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

    $startDelayLabel = New-Object System.Windows.Forms.Label
    $startDelayLabel.Text = "Start delay:"
    $startDelayLabel.Location = New-Object System.Drawing.Point(20, 84)
    $startDelayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $startDelayLabel.Font = $uiFont

    $startDelayInput = New-Object System.Windows.Forms.NumericUpDown
    $startDelayInput.Location = New-Object System.Drawing.Point($valueX, 82)
    $startDelayInput.Size = New-Object System.Drawing.Size(100, 24)
    $startDelayInput.Minimum = 0
    $startDelayInput.Maximum = 86400
    $startDelayInput.Value = [decimal][math]::Min(86400, [math]::Max(0, $StartDelaySeconds))
    $startDelayInput.Font = $uiFont

    $startDelayUnits = New-Object System.Windows.Forms.Label
    $startDelayUnits.Text = "sec"
    $startDelayUnits.Location = New-Object System.Drawing.Point(300, 84)
    $startDelayUnits.Size = New-Object System.Drawing.Size(50, 24)
    $startDelayUnits.Font = $uiFont

    $durationLabel = New-Object System.Windows.Forms.Label
    $durationLabel.Text = "Duration limit:"
    $durationLabel.Location = New-Object System.Drawing.Point(20, 114)
    $durationLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $durationLabel.Font = $uiFont

    $durationInput = New-Object System.Windows.Forms.NumericUpDown
    $durationInput.Location = New-Object System.Drawing.Point($valueX, 112)
    $durationInput.Size = New-Object System.Drawing.Size(100, 24)
    $durationInput.Minimum = 0
    $durationInput.Maximum = 86400
    $durationInput.Value = if ($null -ne $DurationLimitSeconds) { [decimal][math]::Min(86400, [math]::Max(1, $DurationLimitSeconds)) } else { [decimal]0 }
    $durationInput.Font = $uiFont

    $durationUnits = New-Object System.Windows.Forms.Label
    $durationUnits.Text = "sec (0 = off)"
    $durationUnits.Location = New-Object System.Drawing.Point(300, 114)
    $durationUnits.Size = New-Object System.Drawing.Size(130, 24)
    $durationUnits.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $clickLimitLabel = New-Object System.Windows.Forms.Label
    $clickLimitLabel.Text = "Click limit:"
    $clickLimitLabel.Location = New-Object System.Drawing.Point(20, 144)
    $clickLimitLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $clickLimitLabel.Font = $uiFont

    $clickLimitInput = New-Object System.Windows.Forms.NumericUpDown
    $clickLimitInput.Location = New-Object System.Drawing.Point($valueX, 142)
    $clickLimitInput.Size = New-Object System.Drawing.Size(130, 24)
    $clickLimitInput.Minimum = 0
    $clickLimitInput.Maximum = [decimal]2000000000
    $clickLimitInput.Value = if ($null -ne $ClickLimitValue) { [decimal][math]::Min(2000000000, [math]::Max(1, $ClickLimitValue)) } else { [decimal]0 }
    $clickLimitInput.Font = $uiFont

    $clickLimitHint = New-Object System.Windows.Forms.Label
    $clickLimitHint.Text = "(0 = off)"
    $clickLimitHint.Location = New-Object System.Drawing.Point(330, 144)
    $clickLimitHint.Size = New-Object System.Drawing.Size(90, 24)
    $clickLimitHint.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $idleTimeoutLabel = New-Object System.Windows.Forms.Label
    $idleTimeoutLabel.Text = "Idle timeout:"
    $idleTimeoutLabel.Location = New-Object System.Drawing.Point(20, 174)
    $idleTimeoutLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $idleTimeoutLabel.Font = $uiFont

    $idleTimeoutInput = New-Object System.Windows.Forms.NumericUpDown
    $idleTimeoutInput.Location = New-Object System.Drawing.Point($valueX, 172)
    $idleTimeoutInput.Size = New-Object System.Drawing.Size(100, 24)
    $idleTimeoutInput.Minimum = 0
    $idleTimeoutInput.Maximum = 86400
    $idleTimeoutInput.Value = [decimal][math]::Min(86400, [math]::Max(0, $IdleTimeoutSeconds))
    $idleTimeoutInput.Font = $uiFont

    $idleTimeoutUnits = New-Object System.Windows.Forms.Label
    $idleTimeoutUnits.Text = "sec (0 = off)"
    $idleTimeoutUnits.Location = New-Object System.Drawing.Point(300, 174)
    $idleTimeoutUnits.Size = New-Object System.Drawing.Size(130, 24)
    $idleTimeoutUnits.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lifetimeFileLabel = New-Object System.Windows.Forms.Label
    $lifetimeFileLabel.Text = "Lifetime file:"
    $lifetimeFileLabel.Location = New-Object System.Drawing.Point(20, 204)
    $lifetimeFileLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $lifetimeFileLabel.Font = $uiFont

    $lifetimeFileInput = New-Object System.Windows.Forms.TextBox
    $lifetimeFileInput.Location = New-Object System.Drawing.Point($valueX, 202)
    $lifetimeFileInput.Size = New-Object System.Drawing.Size(310, 24)
    $lifetimeFileInput.Text = $script:LifetimeClicksFilePath
    $lifetimeFileInput.Font = New-Object System.Drawing.Font("Consolas", 9)

    $lifetimeBrowseButton = New-Object System.Windows.Forms.Button
    $lifetimeBrowseButton.Text = "Browse..."
    $lifetimeBrowseButton.Location = New-Object System.Drawing.Point(510, 201)
    $lifetimeBrowseButton.Size = New-Object System.Drawing.Size(90, 26)
    $lifetimeBrowseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Status:"
    $statusLabel.Location = New-Object System.Drawing.Point(20, 242)
    $statusLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $statusLabel.Font = $uiFont

    $statusValue = New-Object System.Windows.Forms.Label
    $statusValue.Text = "Stopped"
    $statusValue.Location = New-Object System.Drawing.Point($valueX, 242)
    $statusValue.Size = New-Object System.Drawing.Size(410, 24)
    $statusValue.Font = $uiFont

    $clicksLabel = New-Object System.Windows.Forms.Label
    $clicksLabel.Text = "Clicks:"
    $clicksLabel.Location = New-Object System.Drawing.Point(20, 272)
    $clicksLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $clicksLabel.Font = $uiFont

    $clicksValue = New-Object System.Windows.Forms.Label
    $clicksValue.Text = "0"
    $clicksValue.Location = New-Object System.Drawing.Point($valueX, 272)
    $clicksValue.Size = New-Object System.Drawing.Size(410, 24)
    $clicksValue.Font = $uiFont

    $lifetimeLabel = New-Object System.Windows.Forms.Label
    $lifetimeLabel.Text = "Lifetime clicks:"
    $lifetimeLabel.Location = New-Object System.Drawing.Point(20, 302)
    $lifetimeLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $lifetimeLabel.Font = $uiFont

    $lifetimeValue = New-Object System.Windows.Forms.Label
    $lifetimeValue.Text = $script:LifetimeClicks.ToString()
    $lifetimeValue.Location = New-Object System.Drawing.Point($valueX, 302)
    $lifetimeValue.Size = New-Object System.Drawing.Size(410, 24)
    $lifetimeValue.Font = $uiFont

    $currentDelayLabel = New-Object System.Windows.Forms.Label
    $currentDelayLabel.Text = "Current delay:"
    $currentDelayLabel.Location = New-Object System.Drawing.Point(20, 332)
    $currentDelayLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $currentDelayLabel.Font = $uiFont

    $currentDelayValue = New-Object System.Windows.Forms.Label
    $currentDelayValue.Text = "{0} ms" -f [int]$delayInput.Value
    $currentDelayValue.Location = New-Object System.Drawing.Point($valueX, 332)
    $currentDelayValue.Size = New-Object System.Drawing.Size(410, 24)
    $currentDelayValue.Font = $uiFont

    $lastClickLabel = New-Object System.Windows.Forms.Label
    $lastClickLabel.Text = "Last click:"
    $lastClickLabel.Location = New-Object System.Drawing.Point(20, 362)
    $lastClickLabel.Size = New-Object System.Drawing.Size($labelWidth, 24)
    $lastClickLabel.Font = $uiFont

    $lastClickValue = New-Object System.Windows.Forms.Label
    $lastClickValue.Text = "-"
    $lastClickValue.Location = New-Object System.Drawing.Point($valueX, 362)
    $lastClickValue.Size = New-Object System.Drawing.Size(410, 24)
    $lastClickValue.Font = $uiFont

    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "Start"
    $startButton.Location = New-Object System.Drawing.Point(20, 400)
    $startButton.Size = New-Object System.Drawing.Size(130, 32)
    $startButton.Font = $uiFont

    $stopButton = New-Object System.Windows.Forms.Button
    $stopButton.Text = "Stop"
    $stopButton.Location = New-Object System.Drawing.Point(165, 400)
    $stopButton.Size = New-Object System.Drawing.Size(130, 32)
    $stopButton.Enabled = $false
    $stopButton.Font = $uiFont

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset Count"
    $resetButton.Location = New-Object System.Drawing.Point(310, 400)
    $resetButton.Size = New-Object System.Drawing.Size(130, 32)
    $resetButton.Font = $uiFont

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [int]$delayInput.Value

    $touchInteraction = {
        $state.LastInteractionUtc = [DateTime]::UtcNow
    }

    $setConfigControlsEnabled = {
        param([bool]$Enabled)

        foreach ($control in @(
                $delayInput,
                $jitterCheckBox,
                $startDelayInput,
                $durationInput,
                $clickLimitInput,
                $idleTimeoutInput,
                $lifetimeFileInput,
                $lifetimeBrowseButton
            )) {
            $control.Enabled = $Enabled
        }
    }

    $stopClicker = {
        param([string]$Reason = "Stopped")

        $timer.Stop()
        $state.Phase = 'Stopped'
        $state.RunStartedUtc = $null
        $state.PendingStartUtc = $null

        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        & $setConfigControlsEnabled $true
        $statusValue.Text = $Reason
        Flush-LifetimeClicks
        $state.ClicksSinceFlush = 0
    }

    $lifetimeBrowseButton.Add_Click({
        & $touchInteraction

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = "Select lifetime clicks file"
        $dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $dialog.CheckFileExists = $false
        $dialog.CheckPathExists = $true
        $dialog.FileName = $lifetimeFileInput.Text

        if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $lifetimeFileInput.Text = $dialog.FileName
        }
    })

    $startButton.Add_Click({
        & $touchInteraction

        if ($state.Phase -ne 'Stopped') {
            return
        }

        $requestedLifetimeFile = $lifetimeFileInput.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($requestedLifetimeFile)) {
            [void][System.Windows.Forms.MessageBox]::Show(
                $form,
                "Lifetime file path cannot be empty.",
                "Invalid Lifetime File",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        try {
            $resolvedLifetimePath = Resolve-LifetimeClicksFilePath -PathValue $requestedLifetimeFile
        }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                $form,
                $_.Exception.Message,
                "Invalid Lifetime File",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        if ($resolvedLifetimePath -ne $script:LifetimeClicksFilePath) {
            try {
                Flush-LifetimeClicks
                Initialize-LifetimeClicks -FilePathValue $requestedLifetimeFile
                $lifetimeValue.Text = $script:LifetimeClicks.ToString()
                $lifetimeFileInput.Text = $script:LifetimeClicksFilePath
            }
            catch {
                [void][System.Windows.Forms.MessageBox]::Show(
                    $form,
                    $_.Exception.Message,
                    "Lifetime File Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
        }

        $state.StartDelaySeconds = [int]$startDelayInput.Value
        $durationRaw = [int]$durationInput.Value
        $state.DurationLimitSeconds = if ($durationRaw -gt 0) { [Nullable[int]]$durationRaw } else { $null }
        $clickLimitRaw = [int]$clickLimitInput.Value
        $state.ClickLimitValue = if ($clickLimitRaw -gt 0) { [Nullable[int]]$clickLimitRaw } else { $null }
        $state.IdleTimeoutSeconds = [int]$idleTimeoutInput.Value

        $startButton.Enabled = $false
        $stopButton.Enabled = $true
        & $setConfigControlsEnabled $false
        $state.LastInteractionUtc = [DateTime]::UtcNow

        if ($state.StartDelaySeconds -gt 0) {
            $state.Phase = 'Countdown'
            $state.PendingStartUtc = [DateTime]::UtcNow.AddSeconds($state.StartDelaySeconds)
            $statusValue.Text = "Starting in $($state.StartDelaySeconds)s"
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
    $startDelayInput.Add_ValueChanged({ & $touchInteraction })
    $durationInput.Add_ValueChanged({ & $touchInteraction })
    $clickLimitInput.Add_ValueChanged({ & $touchInteraction })
    $idleTimeoutInput.Add_ValueChanged({ & $touchInteraction })
    $lifetimeFileInput.Add_TextChanged({ & $touchInteraction })

    $timer.Add_Tick({
        $now = [DateTime]::UtcNow

        if ($state.IdleTimeoutSeconds -gt 0 -and ($now - $state.LastInteractionUtc).TotalSeconds -ge $state.IdleTimeoutSeconds) {
            & $stopClicker ("Idle timeout reached ({0}s)" -f $state.IdleTimeoutSeconds)
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

        if ($null -ne $state.DurationLimitSeconds -and ($now - $state.RunStartedUtc).TotalSeconds -ge $state.DurationLimitSeconds) {
            & $stopClicker ("Duration limit reached ({0}s)" -f $state.DurationLimitSeconds)
            return
        }

        if ($null -ne $state.ClickLimitValue -and $state.ClickCount -ge $state.ClickLimitValue) {
            & $stopClicker ("Click limit reached ({0})" -f $state.ClickLimitValue)
            return
        }

        Invoke-LeftClick
        $state.ClickCount++
        $clicksValue.Text = $state.ClickCount.ToString()
        Add-LifetimeClicks -Count 1
        $state.ClicksSinceFlush++
        $lifetimeValue.Text = $script:LifetimeClicks.ToString()
        if ($state.ClicksSinceFlush -ge $flushEveryClicks) {
            Flush-LifetimeClicks
            $state.ClicksSinceFlush = 0
        }

        $baseDelay = [int]$delayInput.Value
        $effectiveDelay = Get-EffectiveDelay -BaseDelay $baseDelay -UseJitter $jitterCheckBox.Checked -Random $state.Random
        $currentDelayValue.Text = "{0} ms" -f $effectiveDelay
        $lastClickValue.Text = Get-Date -Format "HH:mm:ss.fff"
        $timer.Interval = $effectiveDelay

        if ($null -ne $state.ClickLimitValue -and $state.ClickCount -ge $state.ClickLimitValue) {
            & $stopClicker ("Click limit reached ({0})" -f $state.ClickLimitValue)
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

    foreach ($control in @(
            $form,
            $delayInput,
            $jitterCheckBox,
            $startDelayInput,
            $durationInput,
            $clickLimitInput,
            $idleTimeoutInput,
            $lifetimeFileInput,
            $lifetimeBrowseButton,
            $startButton,
            $stopButton,
            $resetButton
        )) {
        $control.Add_MouseMove({
            & $touchInteraction
        })
    }

    $form.Add_FormClosing({
        $timer.Stop()
        Flush-LifetimeClicks
        $timer.Dispose()
        $uiFont.Dispose()
    })

    $form.Controls.AddRange(@(
            $delayLabel,
            $delayInput,
            $delayUnits,
            $jitterCheckBox,
            $startDelayLabel,
            $startDelayInput,
            $startDelayUnits,
            $durationLabel,
            $durationInput,
            $durationUnits,
            $clickLimitLabel,
            $clickLimitInput,
            $clickLimitHint,
            $idleTimeoutLabel,
            $idleTimeoutInput,
            $idleTimeoutUnits,
            $lifetimeFileLabel,
            $lifetimeFileInput,
            $lifetimeBrowseButton,
            $statusLabel,
            $statusValue,
            $clicksLabel,
            $clicksValue,
            $lifetimeLabel,
            $lifetimeValue,
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
Initialize-LifetimeClicks -FilePathValue $LifetimeClicksFile

$useJitter = -not $DisableJitter
$initialDelay = Resolve-InitialDelay -DelayValue $Delay -PromptInConsole ($Mode -eq 'Console')

try {
    switch ($Mode) {
        'Console' {
            Start-ConsoleClicker `
                -InitialDelay $initialDelay `
                -UseJitter $useJitter `
                -StartDelaySeconds $StartDelaySec `
                -DurationLimitSeconds $DurationSec `
                -ClickLimitValue $ClickLimit `
                -IdleTimeoutSeconds $IdleTimeoutSec
        }
        'Gui' {
            Start-GuiClicker `
                -InitialDelay $initialDelay `
                -UseJitter $useJitter `
                -StartDelaySeconds $StartDelaySec `
                -DurationLimitSeconds $DurationSec `
                -ClickLimitValue $ClickLimit `
                -IdleTimeoutSeconds $IdleTimeoutSec
        }
    }
}
finally {
    Flush-LifetimeClicks
}
