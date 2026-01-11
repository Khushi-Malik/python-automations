# Python Automations

A collection of Python scripts for automating tedious file manipulation tasks, saving time and reducing manual effort.

## Overview

This repository contains automation scripts that handle common but time-consuming file operations. The project is designed to be modular and extensible, making it easy to add new automation tasks as needed.

## Current Automations

### 1. Excel to PowerPoint Sync (`main.py`)

Automatically synchronizes data from Excel spreadsheets to PowerPoint presentations while preserving formatting and structure.

**Features:**
- Updates existing PowerPoint table rows with data from Excel
- Automatically adds new rows for tools not found in PowerPoint (optional)
- Handles pagination across multiple slides when tables exceed size limits
- Preserves consistent styling (Calibri 11, configurable bold settings)
- Provides detailed logging of all changes made

**Use Case:**
Perfect for maintaining status tracking presentations, project dashboards, or any recurring reports where data originates in Excel but needs to be presented in PowerPoint format.

**Usage:**
```python
python main.py
```

**Configuration:**
Edit the following variables in `main.py`:
```python
excel_file = "excel.xlsx"           # Path to your Excel file
ppt_file = "ppt.pptx"               # Path to your PowerPoint file
sheet_name = "Procurement AI tracker"  # Excel sheet name
```

**Requirements:**
- Excel columns: `Tool`, `Service\nUse Case`, `Requestor`, `Status`
- PowerPoint columns: `AI Tool`, `Tool Description`, `Requestor`, `Current State`

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)

### Setup

1. Clone this repository:
```bash
git clone <repository-url>
cd python-automations
```

2. Install required dependencies:
```bash
pip install openpyxl python-pptx pandas
```

## Dependencies

- **openpyxl** - Excel file reading and writing
- **python-pptx** - PowerPoint file manipulation
- **pandas** - Data processing and analysis

## Project Structure

```
python-automations/
├── README.md           # This file
├── sync-ppt-with-excel.py            # Excel to PowerPoint sync automation
├── requirements.txt   # Python dependencies (to be added)
└── [future scripts]   # Additional automation tasks
```

## Roadmap

This project will continue to grow with additional automation tasks including:
- PDF manipulation and data extraction
- Batch file renaming and organization
- Document format conversions
- Report generation from multiple data sources
- [Your suggestions here]

## Contributing

This is a personal automation toolkit, but suggestions and improvements are welcome. Feel free to open issues or submit pull requests.

## Best Practices

When adding new automation scripts:
1. Include detailed docstrings explaining functionality
2. Add configuration variables at the top or in a config file
3. Implement error handling and validation
4. Provide progress logging for long-running operations
5. Update this README with usage instructions

## Troubleshooting

### Common Issues

**ImportError: No module named 'openpyxl'**
- Solution: Run `pip install openpyxl python-pptx pandas`

**File not found errors**
- Solution: Ensure file paths in the script match your actual file locations
- Use absolute paths if relative paths aren't working

**PowerPoint formatting issues**
- Solution: Check that your PowerPoint template has the expected column headers
- Verify the table structure matches the required format

## License

This project is provided as-is for personal and educational use.

---

**Note:** This is a living project that will be updated regularly with new automation tasks. Check back for updates!
