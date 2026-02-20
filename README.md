# IBA-Analyzer-GUI

PySide6 GUI for reading **iba PDA .dat files**, replicating the ibaAnalyzer look and feel.

## Requirements

- **Windows** (COM interface)
- **ibaAnalyzer** v8.x (64-bit) installed
- **Python 3.8+** (64-bit)

```bash
pip install PySide6 pywin32 numpy pandas
```

## Usage

```bash
python main.py
```

### Features

- **Signal Tree** — Browse all signals grouped by type (Analog, Digital, Text)
- **Signal Search** — Search signals by wildcard pattern or regex
- **Signal Definitions Table** — Double-click signals to add them to the table
- **Export CSV** — Export selected signals to CSV
- **Export Parquet** — Export selected signals to Parquet (requires `pyarrow`)
- **Export Video** — Extract embedded CaptureCam video to MP4
- **Drag & Drop** — Drag .dat files onto the window to open them

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+O | Open .dat file |
| Delete | Remove selected signal from table |

## Architecture

- `main.py` — GUI application (PySide6)
- `iba_reader.py` — Backend library for reading .dat files via ibaAnalyzer COM

## License

MIT
