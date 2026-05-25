# Inventory Management System

This document provides step-by-step instructions to set up and run the Inventory Management System application on Windows 11.

## Prerequisites

1. **Miniconda Installation**
   - Download Miniconda (included in the project as `Miniconda3-latest-Windows-x86_64.exe`)
   - Run the installer and follow the installation wizard
   - Choose to add Miniconda to your PATH when prompted
   - Restart your computer after installation

2. **Google Cloud Project Setup**
   - Obtain the Google Cloud service account JSON key file (`inventory-managment-465211-7ba8ecdf5815.json`) from your team lead
   - Place the file in the project root directory
   - This file contains sensitive credentials and is not included in the repository for security reasons
   - Do not share or commit this file

## Setup Instructions

1. **Create and Activate Conda Environment**
   ```cmd
   conda create -n streamlit_app python=3.9
   conda activate streamlit_app
   ```

2. **Install Required Packages**
   ```cmd
   pip install -r requirements.txt
   ```

   This will install the following dependencies:
   - streamlit>=1.28.0
   - pandas
   - plotly
   - openpyxl
   - gspread
   - google-auth

## Running the Application

1. **Method 1 - Using Command Prompt**
   ```cmd
   conda activate streamlit_app
   streamlit run app.py
   ```

2. **Method 2 - Using Windows PowerShell**
   ```powershell
   conda activate streamlit_app
   streamlit run app.py
   ```

The application will automatically open in your default web browser. If it doesn't, you can manually open:
- Local URL: http://localhost:8501
- Network URL: http://192.168.x.x:8501 (for accessing from other computers on the same network)

## File Structure
- `app.py` - Main application file containing the Streamlit dashboard code
- `Inventorydata.xlsx` - Excel file containing the inventory data
- `inventory-managment-465211-7ba8ecdf5815.json` - Google Cloud service account credentials (keep secure)
- `requirements.txt` - List of Python package dependencies
- `README.md` - This documentation file

## Data Storage

The application now supports two methods of data storage:

1. **Local Excel File (`Inventorydata.xlsx`)**
   - Traditional local storage method
   - Keep the file in the same directory as `app.py`
   - Close the Excel file before running the application

2. **Google Sheets Integration**
   - Requires the service account JSON key file
   - Enables cloud-based data storage and collaboration
   - Automatically syncs data between local and cloud storage

## Troubleshooting

1. **If Conda Command is Not Recognized**
   - Open Command Prompt as Administrator
   - Run: `where conda`
   - If no path is shown, add Miniconda to your PATH:
     1. Search for "Environment Variables" in Windows Search
     2. Click "Edit the system environment variables"
     3. Click "Environment Variables"
     4. Under "System Variables", find and select "Path"
     5. Click "Edit"
     6. Add the path to your Miniconda installation (typically `C:\Users\YourUsername\miniconda3\Scripts`)
     7. Click "OK" on all windows
     8. Restart Command Prompt

2. **If Excel File Cannot Be Read**
   - Ensure `Inventorydata.xlsx` is in the same directory as `app.py`
   - Close the Excel file if it's open in Excel
   - Make sure you have write permissions in the directory

3. **If Google Sheets Integration Fails**
   - Verify the service account JSON file is in the correct location
   - Check if the file name matches exactly: `inventory-managment-465211-7ba8ecdf5815.json`
   - Ensure the service account has the necessary permissions in Google Cloud Console
   - Check your internet connection

4. **If Streamlit Fails to Start**
   - Ensure you're in the correct conda environment (`conda activate streamlit_app`)
   - Try reinstalling Streamlit: `pip install --force-reinstall streamlit`
   - Check if port 8501 is available (close any other Streamlit applications)

## Support

If you encounter any issues not covered in the troubleshooting section, please:
1. Check if all dependencies are correctly installed: `pip list`
2. Verify you're using the correct Python version: `python --version`
3. Contact the development team with:
   - Screenshots of any error messages
   - Output of `pip list`
   - Your Python version
   - Steps to reproduce the issue
   - Any relevant Google Sheets error messages or access issues

## Security Notes

1. **Service Account Key**
   - Keep the `inventory-managment-465211-7ba8ecdf5815.json` file secure
   - Never commit this file to version control
   - Do not share this file with unauthorized users
   - If the key is compromised, rotate it immediately in Google Cloud Console

2. **Data Access**
   - Regularly review who has access to the Google Sheets
   - Follow the principle of least privilege when granting access
   - Regularly audit access logs in Google Cloud Console 