# BLE Industrial Tool

Desktop BLE host application for scanning, connecting, GATT traffic, and Excel-based register read/write over a custom Modbus-like protocol.

**Author:** F. Y. Wu  
**Open Source Repository:** https://github.com/wfy777/ble
**License:** [GPL-3.0-or-later](LICENSE)

## Features

- BLE scan with device name filter
- Connect, GATT discovery, auto UUID selection
- Excel workbook load with multi-sheet support
- Per-row read (0x03) / write (0x10) dialog

## Requirements

- Python 3.11+
- Windows with Bluetooth (for BLE)
- `pip install bleak openpyxl`

## Run from source

```bash
python ble.py
```

## Build executable

```bash
python -m PyInstaller ble.spec --clean --noconfirm
```

Output: `dist/BLE_Industrial_Tool.exe`

## License

Copyright (C) 2026 F. Y. Wu

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

See [LICENSE](LICENSE) for the full text.

If you distribute a compiled executable, you must also provide the corresponding source code or a written offer to obtain it, as required by the GPL.
