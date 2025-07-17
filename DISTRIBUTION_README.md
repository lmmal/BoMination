# BoMination Application - Sales Team Distribution Package
Version 1.1 - Built on July 17, 2025

## 📦 Package Contents
- **BoMinationApp.exe** (168.8 MB) - The complete application executable
- **QUICK_START.txt** - Quick start instructions for end users
- **deploy.bat** - Optional testing script

## 🚀 Installation Instructions
**No installation required!** This is a portable application.

### For Sales Team Members:
1. Extract the ZIP file to any folder on your computer
2. Double-click `BoMinationApp.exe` to launch
3. If Windows shows a security warning, click "More info" then "Run anyway"

### For IT Departments:
- The application is digitally signed and safe to run
- No system dependencies required - all libraries are bundled
- Runs on Windows 10/11 (64-bit)
- Approximate memory usage: 200-500 MB during operation
- No registry modifications or system changes

## 🔧 System Requirements
- **Operating System:** Windows 10 or Windows 11 (64-bit)
- **Memory:** Minimum 4 GB RAM (8 GB recommended)
- **Storage:** 200 MB free space for temporary files
- **Java:** Recommended but not required (app will prompt if needed)
- **Chrome Browser:** Recommended for price lookup features

## 📋 Application Features
- **PDF Table Extraction:** Automatically extracts Bill of Materials tables from PDF files
- **Price Lookup:** Automatically looks up current pricing for parts
- **Excel Export:** Creates formatted Excel spreadsheets with cost analysis
- **Company-Specific Formatting:** Supports multiple customer formats
- **Manual Review:** Interactive table editing and verification
- **Batch Processing:** Process multiple pages at once

## 🎯 Quick Start Guide

### Step 1: Launch the Application
- Double-click `BoMinationApp.exe`
- The application window will open with a modern interface

### Step 2: Select Your PDF
- Click "Browse" next to "Select BoM PDF File"
- Choose your Bill of Materials PDF file

### Step 3: Specify Page Range
- Enter the page range (e.g., "1-3" for pages 1 through 3)
- Or enter specific pages (e.g., "2,4,6")

### Step 4: Optional Settings
- **Company:** Select if your PDF requires special formatting
- **Table Detection:** Choose detection sensitivity (default: Balanced)
- **Manual Selection:** Enable if automatic detection fails
- **Output Directory:** Choose where to save files (optional)

### Step 5: Run the Process
- Click "Run Automation"
- The application will process your PDF and create Excel files
- Review and confirm the extracted tables when prompted

## 📁 Output Files
The application creates several files in the same folder as your PDF:
- `filename_extracted.xlsx` - Raw extracted table data
- `filename_merged.xlsx` - Cleaned and merged data
- `filename_with_prices.xlsx` - Data with current pricing
- `filename_cost_sheet.xlsx` - Formatted cost analysis

## ⚠️ Troubleshooting

### "Java Not Found" Warning
- Install Java from: https://www.java.com/download/
- Required for PDF table extraction

### "ChromeDriver Not Found" Warning
- Price lookup will be disabled
- Manual pricing can still be added to Excel files

### Application Won't Start
- Try running as Administrator
- Check Windows Defender/antivirus exclusions
- Ensure you have sufficient disk space

### Poor Table Extraction Quality
- Try enabling "Manual Table Area Selection"
- Adjust Table Detection Mode (Conservative/Balanced/Aggressive)
- Ensure PDF has good quality and is not image-only

## 🆘 Support
For technical support:
1. Contact your IT department
2. Include the error message and steps to reproduce
3. Mention this is the "BoMination v1.1" application

## 📧 Distribution Notes
- **File Size:** ~169 MB ZIP file
- **Distribution:** Email attachment or network drive
- **Deployment:** Extract and run - no installation needed
- **Updates:** Replace executable file when new versions are available

## 🔒 Security Information
- Application is built with PyInstaller and digitally verifiable
- No network access except for optional price lookup
- No data is sent externally except for pricing queries
- All processing happens locally on your computer

---
**Built by:** OMNI Engineering Team  
**Build Date:** July 17, 2025  
**Version:** 1.1  
**Package:** BoMination_Sales_Package_v1.1.zip
