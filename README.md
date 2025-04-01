# Installation Guide for Financial Model Analyzer

This guide will walk you through the process of setting up the Financial Model Analyzer on your computer, even if you're not familiar with programming.

## Step 1: Install Python

The Financial Model Analyzer requires Python to run. If you don't have Python installed:

1. Visit the official Python website: https://www.python.org/downloads/
2. Download the latest version of Python for your operating system (Windows, macOS, or Linux)
3. Run the installer
   - **Important for Windows users**: Make sure to check the box that says "Add Python to PATH" during installation
4. Follow the installation prompts to complete the installation

To verify Python is installed correctly, open a command prompt (Windows) or terminal (macOS/Linux) and type:
```
python --version
```

You should see the Python version displayed. If you get an error, try using `python3 --version` instead.

## Step 2: Download the Financial Model Analyzer

1. Create a new folder on your computer where you want to store the application
2. Save the following files to that folder:
   - `financial_analyzer.py` (the main application file)
   - `requirements.txt` (lists required packages)

## Step 3: Install Required Packages

1. Open a command prompt (Windows) or terminal (macOS/Linux)
2. Navigate to the folder where you saved the files
   - On Windows, you can type `cd C:\path\to\your\folder`
   - On macOS/Linux, you can type `cd /path/to/your/folder`
3. Install the required packages by typing:
```
pip install -r requirements.txt
```

If that doesn't work, try:
```
python -m pip install -r requirements.txt
```

Or:
```
python3 -m pip install -r requirements.txt
```

## Step 4: Run the Application

1. In the same command prompt or terminal window, type:
```
python financial_analyzer.py
```

Or if that doesn't work:
```
python3 financial_analyzer.py
```

2. The Financial Model Analyzer application should now launch!

## Troubleshooting

### "Python is not recognized as an internal or external command"
- Make sure Python is added to your PATH
- Try using `python3` instead of `python`
- Restart your computer and try again

### "No module named pandas/openpyxl/etc."
- Make sure you ran the `pip install -r requirements.txt` command successfully
- Try installing packages individually: `pip install pandas openpyxl numpy`

### "Permission denied" errors
- Try running the command prompt or terminal as an administrator (Windows) or using `sudo` (macOS/Linux)

### Application crashes when analyzing a file
- Make sure your Excel file is not corrupted
- Try using a different Excel file to see if the issue persists
- Check that your Excel file contains financial model data in a standard format

## Creating a Desktop Shortcut (Optional)

### Windows:
1. Right-click on your desktop and select New > Shortcut
2. In the location field, type: `python "C:\path\to\your\folder\financial_analyzer.py"`
3. Click Next, name your shortcut "Financial Model Analyzer", and click Finish

### macOS:
1. Open TextEdit and create a new file
2. Write: `#!/bin/bash\ncd /path/to/your/folder/\npython3 financial_analyzer.py`
3. Save as "run_analyzer.command"
4. Open Terminal and type: `chmod +x /path/to/run_analyzer.command`
5. Now you can double-click the file to run the application
