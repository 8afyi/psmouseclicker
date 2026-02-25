# psmouseclicker

Windows auto-clicker script that can run in either console or GUI mode.

## Quick Start

```powershell
# Console mode (default)
.\psmouseclicker.ps1
```

```powershell
# GUI mode
.\psmouseclicker.ps1 -Mode Gui
```

## Modes

- `Console`: prompt-based startup, live status in terminal, `ESC` to stop when key capture is available.
- `Gui`: WinForms window with Start/Stop/Reset controls and live status.

## Parameters

- `-Mode Console|Gui`
  - Choose runtime mode. Default: `Console`.
- `-Delay <int>`
  - Base delay in milliseconds between clicks. If omitted in console mode, script prompts for it.
- `-DisableJitter`
  - Disable the default `+/-10%` randomized jitter.
- `-StartDelaySec <int>`
  - Delay before clicking starts. Default: `0`.
- `-DurationSec <int>`
  - Stop automatically after this many seconds.
- `-ClickLimit <int>`
  - Stop after this many clicks.
- `-IdleTimeoutSec <int>`
  - Stop when no user interaction with the clicker occurs for this many seconds.
  - Set to `0` to disable. Default: `0` (disabled).
- `-LifetimeClicksFile <string>`
  - File used to persist lifetime click totals across runs.
  - Default: `lifetime-clicks.txt` (stored next to the script by default).

## Examples

```powershell
# Console: 250ms delay, no jitter
.\psmouseclicker.ps1 -Mode Console -Delay 250 -DisableJitter
```

```powershell
# GUI: start after 3 seconds, stop after 90 seconds
.\psmouseclicker.ps1 -Mode Gui -Delay 150 -StartDelaySec 3 -DurationSec 90
```

```powershell
# Console: run at most 1000 clicks
.\psmouseclicker.ps1 -ClickLimit 1000
```

```powershell
# GUI: disable idle timeout
.\psmouseclicker.ps1 -Mode Gui -IdleTimeoutSec 0
```

```powershell
# Console: enable 2-minute idle timeout
.\psmouseclicker.ps1 -Delay 30 -IdleTimeoutSec 120
```

```powershell
# Console: store lifetime clicks in a custom file
.\psmouseclicker.ps1 -LifetimeClicksFile C:\temp\psmouseclicker-lifetime.txt
```
