# STORM Monitor - Serial Monitor & Plotter

A feature-rich serial monitor and real-time data plotter built with Python and PyQt5. Designed for embedded development, IoT debugging, and any serial communication workflow.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

---

## Features

- **Serial Communication** — Connect to any COM port with configurable baud rates (9600 to 921600, or custom)
- **Real-Time Serial Monitor** — View incoming serial data with a dark-themed terminal-style display
- **Real-Time Data Plotter** — Plot numeric serial data in real-time with multi-variable support
- **Direct Input Mode** — Send characters one-at-a-time (silent, no echo) for interactive serial devices
- **Command Input** — Send full commands with Enter, with optional echo in the monitor
- **Auto-Reconnect** — Automatically reconnects when a previously connected port becomes available again
- **Data Logging** — Save incoming serial data to `.txt`, `.csv`, or `.bin` files
- **Plot Export** — Export plots as PNG images
- **Dark Theme** — Full dark UI theme with toggleable dark/light plot themes
- **Blinking Cursor & IDLE Indicator** — Visual feedback showing active data reception vs idle state

## Screenshots

> _Coming soon_

## Installation

### From Source

```bash
git clone https://github.com/sunsanpan/STORM-MONITOR.git
cd STORM-MONITOR
pip install -r requirements.txt
python storm_mnonitory_v1.0.py
```

### Pre-built Executable

Download the latest `.exe` from the [`releases/`](releases/) directory — no Python installation required.

## Usage

1. **Select Port** — Choose a COM port from the dropdown or type a custom port name
2. **Set Baud Rate** — Select from common baud rates or type a custom value
3. **Connect** — Click "Connect" to open the serial connection
4. **Monitor** — Incoming data appears in the terminal display
5. **Send Commands** — Type in the command input and press Enter or click "Send"
6. **Direct Input** — Use the direct input field to send individual keystrokes silently
7. **Plotter** — Click "Plotter" to open the real-time data plotting window

### Plotter

The plotter expects numeric data separated by commas, tabs, or spaces. Each value maps to a separate variable line on the graph.

Example serial output that the plotter can visualize:
```
23.5, 45.2, 12.8
24.1, 44.9, 13.0
```

Plotter features:
- Pause/Resume plotting
- Dark/Light theme toggle
- Custom variable names
- Toggle individual variable visibility
- Export plot as PNG

## Project Structure

```
STORM-MONITOR/
├── storm_mnonitory_v1.0.py   # Main application source
├── requirements.txt           # Python dependencies
├── releases/                  # Pre-built executables
│   └── STORM_Monitor_v1.0.exe
├── LICENSE
├── .gitignore
└── README.md
```

## Requirements

- Python 3.8+
- PyQt5
- pyserial
- pyqtgraph
- numpy

## Building the Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name STORM_Monitor_v1.0 storm_mnonitory_v1.0.py
```

The executable will be in the `dist/` folder.

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.

## Author

**sunsanpan**
