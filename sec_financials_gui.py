#!/usr/bin/env python3
"""
GUI version of SEC Financials Tool for non-technical users
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import sys
import os
from pathlib import Path
from sec_financials_tool import create_excel_file

class SECFinancialsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SEC Financials Tool")
        self.root.geometry("550x400")
        self.root.resizable(False, False)
        
        # Try to center the window
        self.center_window()
        
        # Main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="SEC Financials Tool", 
                               font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Subtitle
        subtitle_label = ttk.Label(main_frame, 
                                  text="Generate Excel financial reports from SEC EDGAR",
                                  font=("Arial", 9))
        subtitle_label.grid(row=1, column=0, columnspan=2, pady=(0, 20))
        
        # Ticker input
        ttk.Label(main_frame, text="Company Ticker:", 
                 font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=8)
        self.ticker_var = tk.StringVar()
        ticker_entry = ttk.Entry(main_frame, textvariable=self.ticker_var, width=25, 
                                 font=("Arial", 11))
        ticker_entry.grid(row=2, column=1, sticky=tk.W, pady=8, padx=(10, 0))
        ticker_entry.focus()
        ttk.Label(main_frame, text="(e.g., TSLA, AAPL, MSFT)", 
                 font=("Arial", 8), foreground="gray").grid(row=3, column=1, sticky=tk.W, padx=(10, 0))
        
        # Year input (optional)
        ttk.Label(main_frame, text="Fiscal Year (optional):", 
                 font=("Arial", 10)).grid(row=4, column=0, sticky=tk.W, pady=8)
        self.year_var = tk.StringVar()
        year_entry = ttk.Entry(main_frame, textvariable=self.year_var, width=25, 
                               font=("Arial", 11))
        year_entry.grid(row=4, column=1, sticky=tk.W, pady=8, padx=(10, 0))
        ttk.Label(main_frame, text="(leave blank for latest)", 
                 font=("Arial", 8), foreground="gray").grid(row=5, column=1, sticky=tk.W, padx=(10, 0))
        
        # Email input (optional)
        ttk.Label(main_frame, text="Email (optional):", 
                 font=("Arial", 10)).grid(row=6, column=0, sticky=tk.W, pady=8)
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(main_frame, textvariable=self.email_var, width=30, 
                               font=("Arial", 11))
        email_entry.grid(row=6, column=1, sticky=tk.W, pady=8, padx=(10, 0))
        ttk.Label(main_frame, text="(for SEC API identification)", 
                 font=("Arial", 8), foreground="gray").grid(row=7, column=1, sticky=tk.W, padx=(10, 0))
        
        # Output path (optional)
        ttk.Label(main_frame, text="Save to:", 
                 font=("Arial", 10)).grid(row=8, column=0, sticky=tk.W, pady=8)
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=8, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        self.output_var = tk.StringVar(value="(Auto-generated in current folder)")
        output_label = ttk.Label(output_frame, textvariable=self.output_var, 
                                foreground="gray", font=("Arial", 9), width=30)
        output_label.pack(side=tk.LEFT)
        
        browse_btn = ttk.Button(output_frame, text="Browse...", 
                               command=self.browse_output, width=12)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        self.custom_output_path = None
        
        # Generate button
        self.generate_btn = ttk.Button(main_frame, text="Generate Excel File", 
                                       command=self.generate_file, width=30)
        self.generate_btn.grid(row=9, column=0, columnspan=2, pady=25)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                foreground="blue", font=("Arial", 10))
        status_label.grid(row=10, column=0, columnspan=2, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', length=450)
        self.progress.grid(row=11, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # Bind Enter key to generate button
        self.root.bind('<Return>', lambda e: self.generate_file())
        
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Excel File As"
        )
        if filename:
            self.custom_output_path = filename
            self.output_var.set(os.path.basename(filename))
    
    def generate_file(self):
        ticker = self.ticker_var.get().strip().upper()
        if not ticker:
            messagebox.showerror("Error", "Please enter a company ticker symbol (e.g., TSLA, AAPL, MSFT)")
            return
        
        year = None
        if self.year_var.get().strip():
            try:
                year = int(self.year_var.get().strip())
            except ValueError:
                messagebox.showerror("Error", "Year must be a number (e.g., 2023)")
                return
        
        email = self.email_var.get().strip() or None
        
        # Disable button and show progress
        self.generate_btn.config(state='disabled')
        self.status_var.set("Fetching financial data from SEC...")
        self.progress.start()
        
        # Run in separate thread to avoid freezing GUI
        thread = threading.Thread(target=self._generate_thread, 
                                 args=(ticker, year, email))
        thread.daemon = True
        thread.start()
    
    def _generate_thread(self, ticker, year, email):
        try:
            output_path = create_excel_file(
                ticker=ticker,
                output_path=self.custom_output_path,
                year=year,
                user_email=email
            )
            
            # Update UI in main thread
            self.root.after(0, self._generation_complete, output_path, None)
        except Exception as e:
            self.root.after(0, self._generation_complete, None, str(e))
    
    def _generation_complete(self, output_path, error):
        self.progress.stop()
        self.generate_btn.config(state='normal')
        
        if error:
            self.status_var.set("Error occurred")
            messagebox.showerror("Error", f"Failed to generate file:\n\n{error}\n\nPlease check:\n- Ticker symbol is correct\n- Internet connection is active\n- Try a different year if specified")
        else:
            self.status_var.set("Success!")
            full_path = os.path.abspath(output_path)
            
            # Ask if user wants to open the file
            result = messagebox.askyesno(
                "Success!", 
                f"Excel file created successfully!\n\n"
                f"File: {os.path.basename(output_path)}\n"
                f"Location: {os.path.dirname(full_path)}\n\n"
                f"Would you like to open the file location?",
                icon='question'
            )
            
            if result:
                # Open folder in Finder (Mac) or Explorer (Windows)
                if sys.platform == 'darwin':
                    os.system(f'open -R "{full_path}"')
                elif sys.platform == 'win32':
                    os.system(f'explorer /select,"{full_path}"')
                else:
                    # Linux
                    os.system(f'xdg-open "{os.path.dirname(full_path)}"')

def main():
    root = tk.Tk()
    app = SECFinancialsGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()

