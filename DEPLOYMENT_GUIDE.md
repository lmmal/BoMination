# BoMination Application - Deployment Guide

## Overview
The BoMination application has been successfully packaged into a single executable file for easy distribution to the sales team. The application helps process Farrell PDFs to extract Bill of Materials (BOM) data and map it to cost sheets.

## Files Included
- **`BoMinationApp.exe`** - The main application executable (90.3 MB)
- **`DEPLOYMENT_GUIDE.md`** - This deployment guide
- **`USER_GUIDE.md`** - User instructions for the sales team

## Installation & Distribution

### For IT/Admin:
1. Copy `BoMinationApp.exe` from the `dist/` folder to the target deployment location
2. Ensure the cost sheet template is available at the expected location
3. No additional software installation required - all dependencies are bundled

### For Sales Team:
1. Simply double-click `BoMinationApp.exe` to run the application
2. No installation required - it's a portable executable
3. The application will create necessary temporary files automatically

## System Requirements
- **Operating System**: Windows 10 or later (64-bit)
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Disk Space**: 200MB free space for temporary files and outputs
- **Java**: REQUIRED for PDF table extraction (Java 8 or later)
  - The application will automatically detect Java and prompt for download if missing
  - tabula-py requires Java to extract tables from PDF files
- **Chrome**: Will use system Chrome if available, or bundled ChromeDriver

## File Locations
The application expects and creates files in these locations:
- **Input**: Farrell PDF files (user selects via file dialog)
- **Output**: Processed Excel files in the same directory as the input PDF
- **Cost Sheet Template**: Built into the application
- **ChromeDriver**: Bundled within the executable

## Security Considerations
- The executable is not digitally signed - users may see Windows security warnings
- Windows Defender might flag it initially - this is normal for PyInstaller executables
- Corporate antivirus may require whitelisting the executable

## Troubleshooting

### Common Issues:
1. **Windows Security Warning**: Click "More info" → "Run anyway"
2. **Antivirus Blocking**: Add executable to antivirus whitelist
3. **Application Won't Start**: Ensure Windows 10+ and sufficient disk space
4. **Chrome Issues**: Application includes its own ChromeDriver

### Error Logs:
- Application logs are displayed in the GUI console
- For detailed debugging, run from Command Prompt to see full output

## Features Included
✅ PDF table extraction from Farrell documents  
✅ Automatic header detection and cleaning  
✅ BOM data mapping to cost sheets  
✅ Price lookup from multiple sources  
✅ Excel output with formatted results  
✅ User-friendly GUI interface  
✅ Robust error handling and validation  

## Support
For technical issues or questions, contact the development team with:
- Description of the issue
- Steps to reproduce
- Any error messages displayed
- Sample PDF files (if applicable)

## Version Information
- **Build Date**: Generated with PyInstaller 6.14.1
- **Python Version**: 3.12.6
- **Key Dependencies**: pandas, selenium, ttkbootstrap, tabula-py, openpyxl
- **Package Size**: 90.3 MB (includes all dependencies)

---
*This application was packaged for Omni sales team use. All necessary dependencies and resources are included in the single executable file.*
