# BoMination - Bill of Materials Processing Tool

A Python-based tool for extracting, processing, and managing Bill of Materials (BoM) data from PDF documents.

## Features

- 📄 **PDF Table Extraction**: Extract tables from PDF documents using Tabula
- 🔍 **Interactive Table Selection**: GUI interface to preview and select relevant tables
- 🏢 **Company-Specific Processing**: Specialized formatting for different companies (e.g., Farrell)
- 💰 **Price Lookup**: Automated price matching and validation
- 📊 **Excel Export**: Export processed data to Excel format
- 🔧 **Validation Tools**: Built-in validation and error handling
- 🖥️ **Standalone Executable**: PyInstaller-based executable for easy deployment

## Project Structure

```
BoMination/
├── src/                          # Source code
│   ├── BoMinationApp.py         # Main application entry point
│   ├── extract_bom_tab.py       # PDF table extraction with GUI
│   ├── lookup_price.py          # Price lookup functionality
│   ├── main_pipeline.py         # Main processing pipeline
│   ├── map_cost_sheet.py        # Cost sheet mapping
│   └── validation_utils.py      # Validation utilities
├── Files/                        # Input files and examples
├── SalesTeam_Package/           # Deployment package for sales team
├── requirements.txt             # Python dependencies
├── BoMinationApp.spec          # PyInstaller specification
└── README.md                   # This file
```

## Requirements

- Python 3.8+
- Java Runtime Environment (required for Tabula)
- Chrome/Chromium (for Selenium-based operations)

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/lmmal/BoMination.git
   cd BoMination
   ```

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ensure Java is installed:**
   - Tabula requires Java to be installed on your system
   - Download from [Oracle Java](https://www.oracle.com/java/technologies/downloads/) or [OpenJDK](https://openjdk.org/)

## Usage

### Running the Application

```bash
python src/BoMinationApp.py
```

### Environment Variables

The application uses environment variables for configuration:

- `BOM_PDF_PATH`: Path to the PDF file to process
- `BOM_PAGE_RANGE`: Page range to extract (e.g., "1-5" or "all")
- `BOM_COMPANY`: Company name for specialized processing
- `BOM_OUTPUT_DIRECTORY`: Output directory for processed files

### Example Usage

```bash
# Set environment variables
set BOM_PDF_PATH="C:\path\to\your\document.pdf"
set BOM_PAGE_RANGE="1-3"
set BOM_COMPANY="farrell"
set BOM_OUTPUT_DIRECTORY="C:\output"

# Run the application
python src/BoMinationApp.py
```

## Features in Detail

### PDF Table Extraction
- Uses Tabula library for robust table extraction
- Supports both lattice and stream extraction methods
- Handles complex table structures and formatting

### Interactive Table Selection
- Full-screen GUI for table preview
- Checkbox selection for multiple tables
- Real-time preview of extracted data
- Scroll support for large tables

### Company-Specific Processing
- **Farrell**: Specialized column mapping and formatting
- Extensible architecture for adding new company formats

### Price Lookup
- Automated price matching against databases
- Validation and error reporting
- Support for multiple price sources

## Building Executable

The project includes PyInstaller configuration for creating standalone executables:

```bash
python build_pyinstaller.py
```

## Deployment

The `SalesTeam_Package` folder contains:
- Standalone executable
- Deployment scripts
- User documentation
- Quick start guide

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is proprietary software. All rights reserved.

## Support

For support and questions, please contact the development team.

## Changelog

### Latest Version
- Enhanced table extraction with improved GUI
- Added company-specific processing for Farrell
- Improved error handling and validation
- Added diagnostic tools for troubleshooting

---

*BoMination - Streamlining Bill of Materials processing for efficient operations.*
