# Financial Model Analyzer

A desktop application that extracts and analyzes financial data from Excel files. This tool identifies key assumptions, financial returns, and cash flows from financial models.

## Features

- **Excel File Processing**: Upload and analyze Excel (.xlsx, .xls) financial models
- **Automatic Detection**: Identifies common financial metrics, assumptions, and cash flows
- **Organized Results**: View results in separate tabs for easy navigation
- **Summary Generation**: Provides a comprehensive text summary of all findings

## Getting Started

### Prerequisites

- Python 3.7 or higher
- Required Python packages (install using `pip install -r requirements.txt`):
  - pandas
  - openpyxl
  - numpy

### Installation

1. Clone this repository or download the source code
```
git clone https://github.com/yourusername/financial-model-analyzer.git
cd financial-model-analyzer
```

2. Install required packages
```
pip install -r requirements.txt
```

3. Run the application
```
python financial_analyzer.py
```

## How to Use

1. Launch the application by running `financial_analyzer.py`
2. Click the "Browse" button to select an Excel file containing a financial model
3. Click "Analyze" to process the file
4. Review the extracted information in the various tabs:
   - **Summary**: Text overview of all findings
   - **Assumptions**: Key parameters and variables used in the model
   - **Financial Returns**: NPV, IRR, ROI, and other financial metrics
   - **Cash Flows**: Cash flow projections from the model

## How It Works

The application uses pattern recognition to identify:

- **Assumptions**: By searching for cells containing terms like "assumption," "input," "parameter," or common financial terms
- **Financial Returns**: By identifying typical financial metrics like NPV, IRR, payback period
- **Cash Flows**: By locating cash flow tables and extracting their values

## Limitations

- Works best with well-structured, conventional financial models
- May not correctly identify custom metrics or unconventional formatting
- Designed for personal use and analysis, not for critical financial decision-making

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
