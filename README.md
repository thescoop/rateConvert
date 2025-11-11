# Legal Aid Rate Converter

A GDPR-compliant desktop application for converting time and fee entries between UK Legal Aid case types.

**Developed for:** Woodruff Billing Ltd  
**Author:** Built with assistance from Claude Code  
**Privacy:** All data processing is local - no external transmission

---

## Features

- ✅ **Rate Conversion** - Convert Excel time entries between Legal Aid case types
- ✅ **Auto-Discovery** - Automatically finds rates and Time & Fees files
- ✅ **VBA Preservation** - Maintains Excel macros and formatting
- ✅ **Three-View Preview** - For Review, Input, Output comparison
- ✅ **Modern UI** - Bootstrap themes with graceful fallback
- ✅ **File-Locked Retry** - Safe saving even when Excel file is open
- 🔒 **GDPR Compliant** - All processing local, no data collection

---

## Installation

### Option 1: Using Conda (Recommended)

```bash
# Create environment from file
conda env create -f environment.yml

# Activate environment
conda activate rateConvert

# Run application
python rateConvertGUI.py
```

### Option 2: Using pip

```bash
# Create virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python rateConvertGUI.py
```

### Option 3: Windows Batch File (Easiest)

Simply double-click `_rateConvertGUI.bat` (assumes conda environment exists)

---

## Dependencies

### Required
- **Python 3.8+** (developed with Python 3.13)
- **openpyxl** - Excel file reading/writing with VBA preservation

### Optional (Graceful Fallback)
- **ttkbootstrap** - Modern flatly/darkly bootstrap themes  
  *Fallback: Standard ttk 'clam' theme*
  
- **tkinterdnd2** - Drag-and-drop file upload  
  *Fallback: Click-to-browse file dialog*

**The application runs perfectly without optional dependencies!**

---

## Usage

1. **Launch** the application
2. **Rates File** is auto-discovered (or browse manually)
3. **Drop/Browse** for Time & Fees Excel file
4. **Select** target case type from dropdown
5. **Convert Rates** to process
6. **Review** unmatched entries (if any)
7. **Save Output** when ready

### File Requirements

**Rates Reference File:**
- Excel format (.xlsx or .xlsm)
- Must contain "clearbill" in filename
- First sheet should have case types and activity rates

**Time & Fees Input:**
- Excel format (.xlsx or .xlsm)
- Should contain "Time" in filename (auto-discovery)
- Must have columns: Date, Status, Type, Description, Staff, Hrs/Qty, Amount

---

## Project Structure

```
rateConvert/
├── rateConvertGUI.py          # Main application
├── _rateConvertGUI.bat        # Windows launcher
├── requirements.txt           # pip dependencies
├── environment.yml            # conda environment
├── .gitignore                 # Excludes sensitive Excel/PDF files
└── README.md                  # This file
```

---

## Privacy & Security

- ✅ **Local Processing Only** - No external data transmission
- ✅ **GDPR Compliant** - Privacy-by-design architecture
- ✅ **No Telemetry** - Zero analytics or tracking
- ✅ **Sensitive Files Protected** - .gitignore excludes Excel/PDF
- ✅ **UK Data Protection** - Designed for UK legal sector compliance

**Sensitive client data never leaves your machine.**

---

## Development

### Setting Up for Development

```bash
# Clone repository (private)
git clone https://github.com/thescoop/rateConvert.git
cd rateConvert

# Create environment
conda env create -f environment.yml
conda activate rateConvert

# Run in development mode
python rateConvertGUI.py
```

### Testing

```bash
# Test with ttkbootstrap installed (modern UI)
python rateConvertGUI.py

# Test fallback mode (simulate missing optional deps)
python -c "import sys; sys.modules['ttkbootstrap'] = None; exec(open('rateConvertGUI.py').read())"
```

---

## Troubleshooting

### Application won't start
- Ensure Python 3.8+ is installed
- Check `openpyxl` is installed: `pip install openpyxl`
- Try running from command line to see error messages

### Rates file not found
- Ensure rates file contains "clearbill" in filename
- Place rates file in same directory as application
- Or browse manually using "Change Rates File..." button

### File is locked / Permission denied when saving
- Application automatically retries when file is locked
- Close Excel if file is open
- Check file permissions

### Modern themes not showing
- Install `ttkbootstrap`: `pip install ttkbootstrap`
- Application works fine without it (uses standard theme)

---

## License

**Proprietary - Woodruff Billing Ltd**

This software is confidential and proprietary. Unauthorized copying, distribution, or modification is prohibited.

© 2025 Woodruff Billing Ltd. All rights reserved.

---

## Support

For issues or questions:
1. Check this README
2. Review error messages in Status Log
3. Contact: [Woodruff Billing](https://www.woodruffbilling.co.uk)

---

**Built with Claude Code** - AI-assisted development for modern Python applications.
