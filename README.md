# QuickCell

A minimal read-only `.xlsx` viewer for GNOME/Linux. Built so you can peek at an Excel file someone sent you without firing up LibreOffice.

## Features

- Read-only — the file is never written back
- Sheet tabs (one per worksheet)
- Click/drag to select a cell or a range
- Click a column or row header to select the whole column/row
- Copy (Ctrl+C) copies the selection as tab-separated text, ready to paste into Excel / LibreOffice / Sheets / any text app
- Zoom with Ctrl+scroll, Ctrl+±, or Ctrl+0 to reset
- Keyboard navigation (arrows, Home/End, PgUp/PgDn, Ctrl+Home/End)
- Merged cells respected
- Opens with a filename on the command line

## Installation

### System packages

```bash
# Fedora
sudo dnf install python3-gobject gtk3 cairo-gobject

# Ubuntu/Debian
sudo apt install python3-gi gir1.2-gtk-3.0 python3-cairo

# Arch
sudo pacman -S python-gobject gtk3 cairo
```

### Python deps

```bash
pip install openpyxl
```

## Usage

```bash
python3 quickcell.py                          # empty, then File → Open
python3 quickcell.py /path/to/report.xlsx     # open directly
```

## Controls

| Action | Input |
|---|---|
| Select cell / range | Click / click-drag |
| Extend selection | Shift+click, Shift+arrow |
| Select column / row | Click column / row header |
| Select all | Click top-left corner |
| Navigate | Arrow keys, Home, End, PgUp, PgDn |
| Jump to A1 / last cell | Ctrl+Home / Ctrl+End |
| Copy selection (TSV) | Ctrl+C |
| Open file | Ctrl+O |
| Zoom | Ctrl+scroll, Ctrl++, Ctrl+−, Ctrl+0 |
| Switch sheet | Click tab, or Ctrl+PgUp / Ctrl+PgDn |

## Notes on formulas

QuickCell reads the **cached formula result** stored in the file — the value that Excel (or LibreOffice, Google Sheets, etc.) computed and wrote back the last time it saved. Any spreadsheet file sent to you by someone using a real spreadsheet app will have these cached values, so formulas "just work" for normal inspection.

Files produced purely by scripts (e.g. raw `openpyxl.Workbook().save(...)` without ever being opened by a spreadsheet app) don't have cached values, so formula cells will appear empty.

## License

MIT
