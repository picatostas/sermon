# Serial Monitor

A simple, single-file, cross-platform serial monitor built with tkinter.

## Motivation

I developed most of this as part of another personal project and released it as a standalone
application. The goal is a lean, hackable serial monitor: one readable Python file, the standard
library's GUI toolkit, and pyserial as the only dependency.

## Features

- Runs on Windows, Linux and macOS
- Port list refreshes automatically and shows device descriptions (e.g. `COM7 — CP2102 USB to UART`)
- Any baud rate (type a custom one, e.g. 250000), plus data bits, parity and stop bits
- Flow control: none, RTS/CTS, XON/XOFF, DSR/DTR (DSR/DTR is only enforced on Windows; pyserial ignores it elsewhere)
- Control lines: toggle DTR and RTS, live CTS / DSR / RI / CD indicators, send a break
- Untick DTR before connecting to avoid resetting Arduino-style boards
- Output as text (UTF-8), hex, or escaped bytes; binary data never crashes the reader
- Detects unplugged devices and disconnects cleanly
- Copy output to the clipboard or save it to a file; output is capped at 10,000 lines
- Connection profiles (Preferences → Serial Port Settings)
- Also accepts pyserial URLs such as `socket://host:port`, `rfc2217://host:port` or `loop://`

## Future features

- Themes
- Sending hex and binary data (currently text only)
- Timestamps

... open for suggestions

## Getting started

You need:

- Python 3.8 or later
- Tkinter, which ships with the python.org installers for Windows and macOS.
  On Linux install it with your package manager (e.g. `sudo apt install python3-tk`);
  with Homebrew Python on macOS, run `brew install python-tk`.
- pyserial

```
pip install -r pip-requirements.txt
python serial_mon.py
```

## Preferences

Connection profiles are saved to:

- Windows: `%APPDATA%\sermon\preferences.json`
- macOS: `~/Library/Application Support/sermon/preferences.json`
- Linux: `~/.config/sermon/preferences.json`

On first run the defaults come from `preferences.json` next to the script. Presets from older
versions (`preferences.yaml`) are imported automatically if that file is still present and
PyYAML is installed.

## Systems tested

- Windows 11
- Ubuntu 20.04
- macOS
