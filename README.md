# Monthly Email Exporter

A Python-based GUI application for exporting flagged emails from Microsoft Outlook on a monthly basis. This tool is designed specifically for Rad Law Group. APLC to streamline the process of extracting and organizing flagged emails for billing purposes.

## Features

- **GUI Interface**: User-friendly graphical interface built with tkinter
- **Outlook Integration**: Direct connection to Microsoft Outlook
- **Monthly Filtering**: Automatically filters emails by month
- **Flagged Email Detection**: Identifies and exports flagged emails
- **Configurable Output**: Customizable export folder location
- **Progress Tracking**: Real-time progress indication during export
- **Cross-Platform**: Works on Windows systems with Outlook installed

## Prerequisites

- **Windows Operating System**: This application is designed for Windows
- **Microsoft Outlook**: Must be installed and configured
- **Python 3.7+**: Required for running the application
- **pywin32**: For Outlook integration
- **pyinstaller**: For creating standalone executables

## Installation

### Option 1: Run from Source

1. **Clone or download the repository**
   ```bash
   git clone <repository-url>
   cd monthly-email-exporter
   ```

2. **Download Forest theme files**
   ```bash
   # Clone the Forest-ttk-theme repository
   git clone https://github.com/rdbende/Forest-ttk-theme.git temp-forest
   
   # Copy the light theme files to the root directory
   copy temp-forest\forest-light\* .
   copy temp-forest\forest-light.tcl .
   
   # Clean up temporary directory
   rmdir /s /q temp-forest
   ```

3. **Create a virtual environment (recommended)**
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # On Windows
   ```

4. **Install dependencies**
   ```bash
   pip install -r requirements
   ```

5. **Run the application**
   ```bash
   python gui.py
   ```

### Option 2: Use Pre-built Executable

1. **Download the latest release** from the releases page
2. **Extract the ZIP file**
3. **Run the executable** (`gui.exe`)

## Usage

### First Time Setup

1. **Launch the application**
   - Double-click `gui.exe` or run `python gui.py`

2. **Configure Output Folder**
   - Click "Browse" to select where exported emails will be saved
   - Example: `C:/Users/YourName/Desktop/EmailExports`

3. **Connect to Outlook**
   - Click "Connect to Outlook" to establish connection
   - The application will verify Outlook is installed and accessible

### Monthly Export Process

1. **Select Month**
   - Choose the month you want to export emails from
   - The application automatically detects the current month

2. **Load Flagged Emails**
   - Click "Load Flagged Emails" to scan for flagged emails in the selected month
   - The application will display the count of found emails

3. **Export Emails**
   - Click "Export Emails" to begin the export process
   - A progress window will show the export status
   - Emails will be saved to your configured output folder

## Configuration

The application uses a `config.ini` file to store settings:

```ini
[Folder]
output_folder = C:/Users/YourName/Desktop/EmailExports

[Email]
primary_email = your.email@company.com
```

### Configuration Options

- **output_folder**: Directory where exported emails will be saved
- **primary_email**: Primary email address for the Outlook account

## Project Structure

```
monthly-email-exporter/
├── gui.py                 # Main GUI application
├── config.ini            # Configuration file
├── requirements          # Python dependencies
├── app.ico              # Application icon
├── gui.spec             # PyInstaller specification
├── email-export.ipynb   # Jupyter notebook for development
├── utils/               # Utility modules
│   ├── __init__.py
│   ├── config.py        # Configuration management
│   └── outlook.py       # Outlook integration
├── build/               # Build artifacts (generated)
├── dist/                # Distribution files (generated)
└── .venv/              # Virtual environment (generated)
```

## Development

### Building the Executable

To create a standalone executable:

```bash
pyinstaller gui.spec
```

The executable will be created in the `dist/` directory.

### Development Setup

1. **Install development dependencies**
   ```bash
   pip install -r requirements
   ```

2. **Run in development mode**
   ```bash
   python gui.py
   ```

3. **Use Jupyter notebook for testing**
   - Open `email-export.ipynb` for interactive development

## Troubleshooting

### Common Issues

1. **Outlook Not Found**
   - Ensure Microsoft Outlook is installed
   - Verify Outlook is properly configured with an email account

2. **Permission Errors**
   - Run the application as administrator if needed
   - Check folder permissions for the output directory

3. **No Flagged Emails Found**
   - Verify emails are actually flagged in Outlook
   - Check the selected month contains flagged emails

4. **Export Fails**
   - Ensure the output folder exists and is writable
   - Check available disk space

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Acknowledgments

This project uses the [Forest-ttk-theme](https://github.com/rdbende/Forest-ttk-theme) by [rdbende](https://github.com/rdbende) for the modern GUI styling. The Forest theme provides a beautiful, modern look inspired by MS Excel's design, making the application more visually appealing and user-friendly.

## Support

For technical support or questions, contact the development team at Rad Law Group. APLC.
