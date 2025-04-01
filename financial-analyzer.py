import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook
import openpyxl

class FinancialModelAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Model Analyzer")
        self.root.geometry("800x700")
        self.root.minsize(800, 700)
        
        # Set up the main frame
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Application title
        title_label = ttk.Label(self.main_frame, text="Financial Model Analyzer", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(self.main_frame, 
                               text="Upload your Excel financial model to extract key assumptions and financial returns",
                               wraplength=700)
        desc_label.pack(pady=(0, 20))
        
        # File selection frame
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # File path entry
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=60)
        self.file_path_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        # Browse button
        self.browse_button = ttk.Button(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Analyze button
        self.analyze_button = ttk.Button(self.file_frame, text="Analyze", command=self.analyze_file)
        self.analyze_button.pack(side=tk.LEFT)
        
        # Create a notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.summary_tab = ttk.Frame(self.notebook)
        self.assumptions_tab = ttk.Frame(self.notebook)
        self.returns_tab = ttk.Frame(self.notebook)
        self.cashflow_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.summary_tab, text="Summary")
        self.notebook.add(self.assumptions_tab, text="Assumptions")
        self.notebook.add(self.returns_tab, text="Financial Returns")
        self.notebook.add(self.cashflow_tab, text="Cash Flows")
        
        # Summary Text Area
        self.summary_text = scrolledtext.ScrolledText(self.summary_tab, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create treeviews for other tabs
        self.setup_assumptions_tab()
        self.setup_returns_tab()
        self.setup_cashflows_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready. Please select an Excel file.")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize results dictionary
        self.results = {
            "assumptions": [],
            "financial_returns": {
                "npv": None,
                "irr": None,
                "payback_period": None,
                "roi": None,
                "profit_margin": None
            },
            "cash_flows": []
        }
    
    def setup_assumptions_tab(self):
        # Create frame for treeview
        frame = ttk.Frame(self.assumptions_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create treeview
        columns = ("description", "value")
        self.assumptions_tree = ttk.Treeview(frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)
        
        # Configure scrollbar
        scrollbar.config(command=self.assumptions_tree.yview)
        
        # Set column headings
        self.assumptions_tree.heading("description", text="Description")
        self.assumptions_tree.heading("value", text="Value")
        
        # Set column widths
        self.assumptions_tree.column("description", width=400)
        self.assumptions_tree.column("value", width=200)
        
        # Pack treeview
        self.assumptions_tree.pack(fill=tk.BOTH, expand=True)
    
    def setup_returns_tab(self):
        # Create frame for treeview
        frame = ttk.Frame(self.returns_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create treeview
        columns = ("metric", "value")
        self.returns_tree = ttk.Treeview(frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)
        
        # Configure scrollbar
        scrollbar.config(command=self.returns_tree.yview)
        
        # Set column headings
        self.returns_tree.heading("metric", text="Metric")
        self.returns_tree.heading("value", text="Value")
        
        # Set column widths
        self.returns_tree.column("metric", width=400)
        self.returns_tree.column("value", width=200)
        
        # Pack treeview
        self.returns_tree.pack(fill=tk.BOTH, expand=True)
    
    def setup_cashflows_tab(self):
        # Create frame for treeview
        frame = ttk.Frame(self.cashflow_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create treeview
        columns = ("description", "start_value", "end_value")
        self.cashflow_tree = ttk.Treeview(frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)
        
        # Configure scrollbar
        scrollbar.config(command=self.cashflow_tree.yview)
        
        # Set column headings
        self.cashflow_tree.heading("description", text="Description")
        self.cashflow_tree.heading("start_value", text="Starting Value")
        self.cashflow_tree.heading("end_value", text="Ending Value")
        
        # Set column widths
        self.cashflow_tree.column("description", width=300)
        self.cashflow_tree.column("start_value", width=150)
        self.cashflow_tree.column("end_value", width=150)
        
        # Pack treeview
        self.cashflow_tree.pack(fill=tk.BOTH, expand=True)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"Selected file: {os.path.basename(file_path)}")
    
    def analyze_file(self):
        file_path = self.file_path_var.get()
        
        if not file_path:
            self.status_var.set("Error: Please select a file first")
            return
        
        try:
            self.status_var.set("Analyzing file...")
            self.root.update_idletasks()
            
            # Reset previous results
            self.clear_results()
            
            # Analyze the Excel file
            self.analyze_excel_file(file_path)
            
            # Display results
            self.display_results()
            
            self.status_var.set("Analysis complete")
            
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def clear_results(self):
        # Clear results dictionary
        self.results = {
            "assumptions": [],
            "financial_returns": {
                "npv": None,
                "irr": None,
                "payback_period": None,
                "roi": None,
                "profit_margin": None
            },
            "cash_flows": []
        }
        
        # Clear UI elements
        self.summary_text.delete(1.0, tk.END)
        
        # Clear treeviews
        for item in self.assumptions_tree.get_children():
            self.assumptions_tree.delete(item)
        
        for item in self.returns_tree.get_children():
            self.returns_tree.delete(item)
        
        for item in self.cashflow_tree.get_children():
            self.cashflow_tree.delete(item)
    
    def analyze_excel_file(self, file_path):
        # Load workbook with openpyxl for formula and cell inspection
        workbook = load_workbook(file_path, data_only=True)
        
        # Process each worksheet
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Convert worksheet to a list of lists for easier processing
            data = []
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    row_data.append(cell.value)
                data.append(row_data)
            
            # Find assumptions
            self.find_assumptions(data)
            
            # Find financial returns
            self.find_financial_returns(data)
            
            # Find cash flows
            self.find_cash_flows(data)
    
    def find_assumptions(self, data):
        assumption_keywords = ['assumption', 'input', 'parameter', 'variable']
        
        for row_idx, row in enumerate(data):
            if not row or len(row) == 0:
                continue
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                
                cell_str = str(cell).lower() if isinstance(cell, (str, int, float)) else ""
                
                # Check if this cell contains assumption keywords
                if any(keyword in cell_str for keyword in assumption_keywords):
                    # Look for value in adjacent cells
                    value = None
                    
                    # Check right cell first
                    if col_idx + 1 < len(row) and row[col_idx + 1] is not None:
                        value = row[col_idx + 1]
                    # Then check cell below
                    elif row_idx + 1 < len(data) and col_idx < len(data[row_idx + 1]) and data[row_idx + 1][col_idx] is not None:
                        value = data[row_idx + 1][col_idx]
                    
                    if value is not None:
                        self.results["assumptions"].append({
                            "description": str(cell),
                            "value": value
                        })
        
        # Also look for common financial model assumptions
        common_assumptions = [
            'discount rate', 'growth rate', 'tax rate', 'inflation', 
            'capex', 'opex', 'revenue', 'cost'
        ]
        
        for row_idx, row in enumerate(data):
            if not row or len(row) == 0:
                continue
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                
                cell_str = str(cell).lower() if isinstance(cell, (str, int, float)) else ""
                
                for assumption in common_assumptions:
                    if assumption in cell_str:
                        # Look for value in adjacent cells
                        value = None
                        
                        # Check right cell first
                        if col_idx + 1 < len(row) and row[col_idx + 1] is not None:
                            value = row[col_idx + 1]
                        # Then check cell below
                        elif row_idx + 1 < len(data) and col_idx < len(data[row_idx + 1]) and data[row_idx + 1][col_idx] is not None:
                            value = data[row_idx + 1][col_idx]
                        
                        if value is not None:
                            # Avoid duplicates
                            if not any(a["description"] == str(cell) for a in self.results["assumptions"]):
                                self.results["assumptions"].append({
                                    "description": str(cell),
                                    "value": value
                                })
                        
                        break
    
    def find_financial_returns(self, data):
        return_indicators = {
            'npv': ['npv', 'net present value'],
            'irr': ['irr', 'internal rate of return'],
            'payback_period': ['payback', 'payback period'],
            'roi': ['roi', 'return on investment'],
            'profit_margin': ['profit margin', 'margin']
        }
        
        for row_idx, row in enumerate(data):
            if not row or len(row) == 0:
                continue
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                
                cell_str = str(cell).lower() if isinstance(cell, (str, int, float)) else ""
                
                # Check each financial indicator
                for key, keywords in return_indicators.items():
                    if any(keyword in cell_str for keyword in keywords):
                        # Look for value in adjacent cells
                        value = None
                        
                        # Check right cell first
                        if col_idx + 1 < len(row) and row[col_idx + 1] is not None:
                            value = row[col_idx + 1]
                        # Then check cell below
                        elif row_idx + 1 < len(data) and col_idx < len(data[row_idx + 1]) and data[row_idx + 1][col_idx] is not None:
                            value = data[row_idx + 1][col_idx]
                        
                        if value is not None and self.results["financial_returns"][key] is None:
                            self.results["financial_returns"][key] = {
                                "label": str(cell),
                                "value": value
                            }
                        
                        break
    
    def find_cash_flows(self, data):
        # Look for cash flow tables
        cashflow_keywords = ['cash flow', 'cashflow', 'cash flows', 'cf']
        
        # Find potential cash flow tables
        for row_idx, row in enumerate(data):
            if not row or len(row) == 0:
                continue
            
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                
                cell_str = str(cell).lower() if isinstance(cell, (str, int, float)) else ""
                
                if any(keyword in cell_str for keyword in cashflow_keywords):
                    # Found potential cash flow table, try to extract it
                    cashflows = self.extract_cash_flow_table(data, row_idx, col_idx)
                    if cashflows:
                        self.results["cash_flows"].extend(cashflows)
                        return  # Just use the first cash flow table found for simplicity
    
    def extract_cash_flow_table(self, data, start_row, start_col):
        cashflows = []
        header_row = None
        
        # Look for header row (years or periods)
        for i in range(start_row, min(start_row + 5, len(data))):
            row = data[i]
            if not row:
                continue
            
            period_count = 0
            for j in range(1, len(row)):
                cell = row[j]
                if cell is not None and (
                    isinstance(cell, (int, float)) or 
                    (isinstance(cell, str) and re.search(r'year|period|yr|\d+', cell, re.IGNORECASE))
                ):
                    period_count += 1
            
            if period_count >= 3:  # At least 3 periods to consider it a cashflow table
                header_row = i
                break
        
        if header_row is None:
            return cashflows
        
        # Extract cash flow data
        current_row = header_row + 1
        while current_row < len(data) and current_row < header_row + 20:
            row = data[current_row]
            if not row or len(row) == 0:
                current_row += 1
                continue
            
            first_cell = row[0]
            if first_cell is None:
                current_row += 1
                continue
            
            first_cell_str = str(first_cell).lower() if isinstance(first_cell, (str, int, float)) else ""
            
            if ('cash flow' in first_cell_str or 
                'net' in first_cell_str or 
                'total' in first_cell_str):
                
                periods = []
                for i in range(1, len(row)):
                    if i < len(row) and row[i] is not None:
                        periods.append({
                            "period": i,
                            "value": row[i]
                        })
                
                if periods:
                    cashflows.append({
                        "label": str(first_cell),
                        "periods": periods
                    })
            
            current_row += 1
        
        return cashflows
    
    def display_results(self):
        # Generate and display summary
        summary = self.generate_summary()
        self.summary_text.insert(tk.END, summary)
        
        # Display assumptions
        for assumption in self.results["assumptions"]:
            self.assumptions_tree.insert("", tk.END, values=(
                assumption["description"], 
                assumption["value"]
            ))
        
        # Display financial returns
        for key, value in self.results["financial_returns"].items():
            if value is not None:
                # Format label
                formatted_label = key.replace('_', ' ').title()
                
                self.returns_tree.insert("", tk.END, values=(
                    value["label"], 
                    value["value"]
                ))
        
        # Display cash flows
        for cf in self.results["cash_flows"]:
            if not cf["periods"]:
                continue
                
            first = cf["periods"][0]
            last = cf["periods"][-1]
            
            self.cashflow_tree.insert("", tk.END, values=(
                cf["label"],
                first["value"],
                last["value"]
            ))
    
    def generate_summary(self):
        summary = "Financial Model Analysis Summary:\n\n"
        
        # Assumptions
        if self.results["assumptions"]:
            summary += "Key Assumptions:\n"
            for assumption in self.results["assumptions"]:
                summary += f"- {assumption['description']}: {assumption['value']}\n"
            summary += "\n"
        else:
            summary += "No clear assumptions were identified in this model.\n\n"
        
        # Financial Returns
        summary += "Financial Returns:\n"
        has_returns = False
        
        for key, value in self.results["financial_returns"].items():
            if value is not None:
                has_returns = True
                
                # Format label
                formatted_label = key.replace('_', ' ').title()
                
                summary += f"- {value['label'] or formatted_label}: {value['value']}\n"
        
        if not has_returns:
            summary += "No clear financial return indicators were identified in this model.\n"
        
        # Cash Flows
        if self.results["cash_flows"]:
            summary += "\nCash Flow Summary:\n"
            for cf in self.results["cash_flows"]:
                summary += f"- {cf['label']}: "
                
                # Get first and last period for summary
                if cf["periods"]:
                    first = cf["periods"][0]
                    last = cf["periods"][-1]
                    
                    summary += f"Starts at {first['value']} and ends at {last['value']}\n"
                else:
                    summary += "No period data available\n"
        
        return summary


if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialModelAnalyzer(root)
    root.mainloop()
