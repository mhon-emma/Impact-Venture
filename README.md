# Gemini AI Financial Model Analyzer

An advanced Python application for analyzing Excel financial models using Google's Gemini AI. This tool extracts assumptions, financial returns, and cash flow data from Excel financial models of any structure or format.

## Features

- **Gemini AI-Powered Analysis**: Uses Google's Gemini model to analyze financial models intelligently
- **Structure-Agnostic**: Works with a wide variety of Excel financial model formats without relying on keyword search
- **Key Data Extraction**:
  - Identifies key assumptions and input parameters
  - Extracts financial returns (NPV, IRR, ROI, profit margins, etc.)
  - Analyzes cash flow projections
- **Modern UI**: Clean interface with tabbed results display
- **Progress Tracking**: Visual feedback during analysis process
- **Large File Support**: Handles large Excel files with intelligent data extraction

## Requirements

- Python 3.7+
- Google Gemini API key (available from [Google AI Studio](https://makersuite.google.com/app/apikey))
- Required Python packages (included in requirements.txt):
  - tkinter
  - pandas
  - numpy
  - openpyxl
  - requests

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/gemini-financial-analyzer.git
cd gemini-financial-analyzer
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python financial_analyzer.py
```

## Usage

1. Launch the application
2. Enter your Gemini API key in the Settings dialog (first-time setup)
3. Click "Browse" to select an Excel file containing a financial model
4. Click "Analyze with Gemini" to process the file
5. View the extracted information in the corresponding tabs:
   - Summary: Overview of the financial model
   - Assumptions: Key input parameters
   - Financial Returns: NPV, IRR, ROI and other metrics
   - Cash Flows: Cash flow projections summary

## How It Works

The application extracts and processes Excel data in several steps:

1. **File Processing**:
   - The app reads the Excel file using openpyxl
   - It extracts data from each sheet (limiting to 100 rows per sheet for large files)
   - All data is converted to a format suitable for AI analysis

2. **AI Analysis**:
   - The extracted data is sent to Google's Gemini API
   - The AI analyzes the data structure to identify financial components
   - Results are returned in structured JSON format

3. **Results Display**:
   - The application parses the AI's analysis
   - Results are presented in a clean, organized interface with different tabs

## API Key Security

- Your Gemini API key is stored securely in a local configuration file at `~/.financial_analyzer/config.ini`
- Keys are never transmitted except to the Gemini API for authorization
- Option to show/hide the API key in the settings dialog

## Benefits Over Traditional Methods

- **No Keyword Dependencies**: Traditional financial analysis tools rely on specific keywords or cell locations, making them brittle when Excel templates change. This tool adapts to any format.
- **Contextual Understanding**: Gemini AI understands the broader context of financial data, not just isolated cells.
- **Comprehensive Analysis**: The AI can identify relationships between data points that rule-based systems might miss.

## Limitations

- Requires an internet connection for AI processing
- Analysis quality depends on the Gemini API's capabilities
- Very complex or unusual financial models may not be fully understood

## Future Enhancements

- Support for more file formats (CSV, Google Sheets)
- Offline mode with local model option
- Custom templates for specific financial model types
- Export capabilities for analysis results
- Batch processing of multiple files

## License

This project is licensed under the MIT License - see the LICENSE file for details.