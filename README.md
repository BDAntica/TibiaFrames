# TibiaFrames

A desktop app for browsing, organizing, and viewing Tibia game screenshots. Runs on Windows.

## Features

- **Tree browser** — screenshots organized by character, category, and date
- **Image viewer** — pan and zoom with mouse, keyboard navigation between screenshots
- **Copy to clipboard** — one-click copy of any screenshot (fast, no delay)
- **Statistics panel** — activity breakdown by hour, day of week, month, and year
- **Open folder** — jump straight to the folder containing any screenshot
- **Dark/light theme** — toggleable from the toolbar
- **Persistent window** — remembers size and position between sessions
- **Auto-loads** default Tibia screenshot directory on startup

## Supported screenshot categories

Achievements, Bestiary entries, Boss defeats, Deaths (PvE/PvP), Gift of Life, Level ups, Skill ups, Player kills, Loot, Hotkeys, and more — all parsed directly from Tibia's filename format.

## Requirements

- Windows 10 or later
- [pywin32](https://pypi.org/project/pywin32/) for fast clipboard support (optional — falls back to PowerShell if not installed)

## Running from source

```
pip install pillow pywin32
python tibiaframes_v1_2_4.pyw
```

## Running the compiled EXE

Download `TibiaFrames.exe` from the `dist/` folder and run it directly — no installation needed.

## Building the EXE yourself

```
build.bat
```

Requires Python, Pillow, and PyInstaller. The script cleans previous build artifacts and produces `dist/TibiaFrames.exe`.

## Screenshot directory

TibiaFrames looks for screenshots in the default Tibia path on startup:

```
C:\Users\<you>\Documents\Tibia\Screenshots
```

You can also browse to any folder manually from the toolbar.

## Version

v1.2.4
