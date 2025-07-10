# BoMination User Guide

## Quick Start
1. **Launch**: Double-click `BoMinationApp.exe`
2. **Select PDF**: Click "Browse" to choose your Farrell PDF file
3. **Extract Tables**: Click "Extract Tables" to analyze the PDF
4. **Select Tables**: Check the boxes for tables you want to process
5. **Run Pipeline**: Click "Run Main Pipeline" to generate the cost sheet

## Step-by-Step Instructions

### 1. Starting the Application
- Locate the `BoMinationApp.exe` file
- Double-click to launch (you may see a Windows security warning - click "Run anyway")
- The application window will open with a clean interface

### 2. Loading a PDF File
- Click the **"Browse"** button next to "Select PDF File"
- Navigate to your Farrell PDF document
- Select the file and click "Open"
- The file path will appear in the text field

### 3. Extracting Tables
- Click **"Extract Tables"** button
- The application will:
  - Analyze the PDF structure
  - Identify tables automatically
  - Display found tables in the preview area
- This process may take 30-60 seconds depending on PDF size

### 4. Reviewing and Selecting Tables
- **Table Preview**: Review the extracted tables in the display area
- **Table Selection**: Check the boxes next to tables that contain BOM data
- **Quality Check**: Ensure tables have proper headers (Part Number, Description, Quantity, etc.)

### 5. Processing the Data
- Click **"Run Main Pipeline"** after selecting tables
- The application will:
  - Clean and standardize the data
  - Look up pricing information
  - Map data to the cost sheet template
  - Generate an Excel output file

### 6. Reviewing Results
- The output Excel file will be saved in the same folder as your input PDF
- File name format: `[Original_PDF_Name]_processed_[timestamp].xlsx`
- Open the Excel file to review the cost sheet with populated BOM data

## Understanding the Interface

### Main Sections:
- **File Selection**: Top area for choosing your PDF file
- **Table Extraction**: Middle section with extract button and console output
- **Table Preview**: Large area showing extracted tables with selection checkboxes
- **Processing**: Bottom section with pipeline button and status updates

### Console Output:
- Shows real-time progress and status messages
- Debug information for troubleshooting
- Error messages if issues occur

## Best Practices

### PDF Quality:
- ✅ Use clear, high-quality PDF scans
- ✅ Ensure tables are properly aligned
- ✅ Text should be selectable (not just images)
- ❌ Avoid heavily skewed or rotated documents

### Table Selection:
- ✅ Select tables with complete BOM data
- ✅ Look for tables with Part Number, Description, Quantity columns
- ❌ Avoid selecting summary tables or partial data

### Data Validation:
- ✅ Review the extracted data before processing
- ✅ Check that headers are correctly identified
- ✅ Verify part numbers and quantities look accurate

## Common Issues & Solutions

### Issue: "No tables found"
**Solution**: 
- Check if PDF has selectable text (not just scanned images)
- Try a different PDF or higher quality scan
- Ensure the document actually contains tabular data

### Issue: "Tables are empty or garbled"
**Solution**:
- PDF quality may be poor
- Tables might be in image format
- Try re-scanning the document at higher resolution

### Issue: "Price lookup failed"
**Solution**:
- Check internet connection
- Some part numbers may not be found in pricing databases
- Manual price entry may be required in the output Excel

### Issue: "Application crashes or freezes"
**Solution**:
- Restart the application
- Try processing smaller PDF files first
- Ensure sufficient disk space and memory

### Issue: "Chrome/Selenium errors"
**Solution**:
- Application includes its own ChromeDriver
- If issues persist, try restarting the application
- Check that no other browser automation is running

## Output File Structure

The generated Excel file contains:
- **Original BOM Data**: Cleaned and standardized
- **Pricing Information**: Looked up from various sources
- **Cost Calculations**: Extended totals and summaries
- **Part Details**: Descriptions, manufacturers, specifications

## Tips for Success

1. **Start Small**: Test with simple PDFs first
2. **Quality Matters**: Better PDF quality = better results
3. **Review Data**: Always verify the extracted information
4. **Save Originals**: Keep backup copies of your source PDFs
5. **Check Results**: Review the output Excel before using in quotes

## Important First-Time Setup

### Windows Security Warning (Expected!)
When you first run the application, Windows will show a security warning:
1. Click **"More info"** 
2. Then click **"Run anyway"**
3. This is normal for applications like this - it's not a virus!

### Antivirus Software
Some antivirus programs may flag the application:
- This is a "false positive" - the app is safe
- If blocked, contact IT to add `BoMinationApp.exe` to the whitelist
- This is common with packaged Python applications

### Internet Connection Required
The application needs internet access for:
- Looking up part pricing information
- Web-based data validation
- Ensure you're not behind a restrictive firewall

## Getting Help

If you encounter issues:
1. Check the console output for error messages
2. Try the troubleshooting steps above
3. Contact your IT support with:
   - Screenshots of any errors
   - The problematic PDF file (if possible)
   - Description of what you were trying to do

---
*BoMination Application - Making BOM processing simple and efficient for the Omni sales team*
