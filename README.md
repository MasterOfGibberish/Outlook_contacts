# Outlook Contact Exporter Add-in

This add-in adds a "Save Contacts" button to Microsoft Outlook that lets you extract contact information from your emails and save it to an Excel file.

## Features

✅ Adds a "Contact Tools" tab in Outlook with a "Save Contacts" button
✅ Scans Sent Items and other folders for all contacts
✅ Extracts names and email addresses
✅ Separates first and last names
✅ Saves the data to a timestamped Excel file on your Desktop
✅ Removes duplicate entries
✅ Works with multiple Outlook versions (2013, 2016, 2019, 365)

## Screenshot

![Outlook Tab](https://raw.githubusercontent.com/MasterOfGibberish/Outlook_contacts/main/screenshots/outlook_tab.png)
*Contact Tools tab in Outlook*

## Requirements

- Microsoft Windows (Windows 10 or 11 recommended)
- Microsoft Outlook (Desktop version - 2013 or newer)
- Python 3.7 or higher (will be installed automatically if missing)

## Installation from GitHub

### Quick Install (Recommended)

1. [Download this repository as ZIP](https://github.com/MasterOfGibberish/Outlook_contacts/archive/refs/heads/main.zip)
2. Extract the ZIP file to a folder on your computer
3. Right-click on `install_addin.py` and select **"Run as administrator"**
4. Choose option 1 (Install Add-in)
5. Restart Microsoft Outlook

### Manual Installation

If the quick install doesn't work:

1. Make sure Python is installed on your computer (Python 3.7 or higher)
2. Open a command prompt as administrator
3. Navigate to the folder where you extracted the files
4. Run: `pip install pywin32 pandas openpyxl`
5. Run: `python install_addin.py`
6. Choose option 1 (Install Add-in)
7. Restart Microsoft Outlook

## Usage

1. Open Microsoft Outlook
2. Look for the new "Contact Tools" tab in the ribbon
3. Click the "Save Contacts" button
4. The add-in will scan your email folders and extract all contacts
5. An Excel file named "outlook_contacts_YYYYMMDD_HHMMSS.xlsx" will be saved to your Desktop
6. A confirmation message will appear in your Outlook

## Troubleshooting

### Common Issues

- **Missing "Contact Tools" tab:** Restart Outlook completely. If still missing, re-run the installer.
- **"Not authorized" error:** Make sure you ran the installer as administrator.
- **Installation fails:** Try the manual installation method above.
- **Excel file not created:** Check your Desktop for "contact_export_result.txt" or look in %TEMP% folder.
- **Add-in crashes:** Look for "outlook_contact_exporter_error.txt" on your Desktop.
- **"No such file or directory" error:** Make sure you're running the installer from the folder where it's located, not copying it elsewhere.

### Logs

The add-in creates logs at: `%LOCALAPPDATA%\OutlookContactExporter\addin_log.txt`

These logs can help diagnose issues if you need support.

## Uninstallation

1. Right-click on `install_addin.py` and select **"Run as administrator"**
2. Choose option 2 (Uninstall Add-in)
3. Restart Microsoft Outlook

## Advanced Options

For power users, you can run the extraction directly without the Outlook add-in:

1. Open a command prompt
2. Navigate to the folder with the files
3. Run: `python extract_contacts.py`

This will open a GUI that performs the same extraction without needing the add-in installed.

## Support

If you encounter any issues:

1. Check the [GitHub Issues](https://github.com/MasterOfGibberish/Outlook_contacts/issues) for known problems
2. Submit a new issue with:
   - Your Windows version
   - Your Outlook version
   - Any error messages
   - Log file contents

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 