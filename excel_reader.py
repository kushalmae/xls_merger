#!/usr/bin/env python3
"""
Excel Multi-Sheet Reader

This script provides comprehensive functionality to read and process
Excel files with multiple sheets. It can handle the data.xlsx file
with its definitions, monitors, and conditions sheets.

Usage: python excel_reader.py
"""

import pandas as pd
import json
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

def log_message(message: str):
    """Print timestamped log message"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

class ExcelMultiSheetReader:
    """Class to handle reading and processing multi-sheet Excel files"""
    
    def __init__(self, file_path: str):
        """Initialize with Excel file path"""
        self.file_path = file_path
        self.sheets_data = {}
        self.xl_file = None
        
    def load_file(self) -> bool:
        """Load the Excel file and check available sheets"""
        try:
            if not os.path.exists(self.file_path):
                log_message(f"ERROR: File '{self.file_path}' not found!")
                return False
            
            self.xl_file = pd.ExcelFile(self.file_path)
            log_message(f"Successfully loaded: {self.file_path}")
            log_message(f"Available sheets: {self.xl_file.sheet_names}")
            return True
            
        except Exception as e:
            log_message(f"ERROR: Failed to load file: {e}")
            return False
    
    def read_all_sheets(self) -> Dict[str, pd.DataFrame]:
        """Read all sheets and return as dictionary of DataFrames"""
        if not self.xl_file:
            log_message("ERROR: File not loaded. Call load_file() first.")
            return {}
        
        try:
            for sheet_name in self.xl_file.sheet_names:
                log_message(f"Reading sheet: {sheet_name}")
                df = pd.read_excel(self.xl_file, sheet_name=sheet_name)
                self.sheets_data[sheet_name] = df
                log_message(f"  - Shape: {df.shape}")
                log_message(f"  - Columns: {list(df.columns)}")
            
            return self.sheets_data
            
        except Exception as e:
            log_message(f"ERROR: Failed to read sheets: {e}")
            return {}
    
    def read_specific_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Read a specific sheet by name"""
        if not self.xl_file:
            log_message("ERROR: File not loaded. Call load_file() first.")
            return None
        
        if sheet_name not in self.xl_file.sheet_names:
            log_message(f"ERROR: Sheet '{sheet_name}' not found!")
            log_message(f"Available sheets: {self.xl_file.sheet_names}")
            return None
        
        try:
            df = pd.read_excel(self.xl_file, sheet_name=sheet_name)
            log_message(f"Successfully read sheet '{sheet_name}'")
            log_message(f"  - Shape: {df.shape}")
            log_message(f"  - Columns: {list(df.columns)}")
            return df
            
        except Exception as e:
            log_message(f"ERROR: Failed to read sheet '{sheet_name}': {e}")
            return None
    
    def display_sheet_summary(self):
        """Display a summary of all sheets"""
        if not self.sheets_data:
            log_message("No sheets loaded. Call read_all_sheets() first.")
            return
        
        log_message("\n" + "="*60)
        log_message("SHEETS SUMMARY")
        log_message("="*60)
        
        for sheet_name, df in self.sheets_data.items():
            log_message(f"\nSheet: {sheet_name}")
            log_message("-" * 40)
            log_message(f"Rows: {len(df)}")
            log_message(f"Columns: {list(df.columns)}")
            
            # Show first few rows
            log_message("First few rows:")
            for i, row in df.head(3).iterrows():
                log_message(f"  Row {i}: {dict(row)}")
            
            if len(df) > 3:
                log_message(f"  ... and {len(df) - 3} more rows")
    
    def get_definitions_mapping(self) -> Optional[Dict[str, str]]:
        """Get the FDIR to ID mapping from definitions sheet"""
        if 'definitions' not in self.sheets_data:
            log_message("Definitions sheet not loaded")
            return None
        
        df = self.sheets_data['definitions']
        if 'FDIRs' in df.columns and 'id' in df.columns:
            mapping = dict(zip(df['FDIRs'], df['id']))
            log_message("Definitions mapping:")
            for fdir, id_val in mapping.items():
                log_message(f"  {fdir} -> {id_val}")
            return mapping
        else:
            log_message("ERROR: Expected columns 'FDIRs' and 'id' not found in definitions sheet")
            return None
    
    def get_monitors_for_id(self, target_id: str) -> Optional[pd.DataFrame]:
        """Get all monitors for a specific ID"""
        if 'monitors' not in self.sheets_data:
            log_message("Monitors sheet not loaded")
            return None
        
        df = self.sheets_data['monitors']
        if 'id' not in df.columns:
            log_message("ERROR: 'id' column not found in monitors sheet")
            return None
        
        filtered_df = df[df['id'] == target_id]
        log_message(f"Found {len(filtered_df)} monitors for ID '{target_id}'")
        return filtered_df
    
    def get_conditions_for_id(self, target_id: str) -> Optional[pd.DataFrame]:
        """Get all conditions for a specific ID"""
        if 'conditions' not in self.sheets_data:
            log_message("Conditions sheet not loaded")
            return None
        
        df = self.sheets_data['conditions']
        if 'id' not in df.columns:
            log_message("ERROR: 'id' column not found in conditions sheet")
            return None
        
        # Filter out NaN rows and get conditions for target ID
        filtered_df = df[df['id'] == target_id].dropna(subset=['condition_mons'])
        log_message(f"Found {len(filtered_df)} conditions for ID '{target_id}'")
        return filtered_df
    
    def export_to_json(self, output_file: str) -> bool:
        """Export all sheets data to JSON file"""
        if not self.sheets_data:
            log_message("No data to export")
            return False
        
        try:
            # Convert DataFrames to dictionaries for JSON serialization
            json_data = {}
            for sheet_name, df in self.sheets_data.items():
                json_data[sheet_name] = df.to_dict('records')
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            log_message(f"✓ Exported all sheets to: {output_file}")
            return True
            
        except Exception as e:
            log_message(f"✗ Failed to export to JSON: {e}")
            return False
    
    def export_sheet_to_csv(self, sheet_name: str, output_file: str) -> bool:
        """Export a specific sheet to CSV"""
        if sheet_name not in self.sheets_data:
            log_message(f"Sheet '{sheet_name}' not loaded")
            return False
        
        try:
            self.sheets_data[sheet_name].to_csv(output_file, index=False)
            log_message(f"✓ Exported sheet '{sheet_name}' to: {output_file}")
            return True
            
        except Exception as e:
            log_message(f"✗ Failed to export sheet to CSV: {e}")
            return False


def interactive_mode():
    """Interactive mode for exploring Excel files"""
    print("\n" + "="*60)
    print("EXCEL MULTI-SHEET READER - INTERACTIVE MODE")
    print("="*60)
    
    # Get file path
    file_path = input("\nEnter Excel file path (or press Enter for 'inputs/data.xlsx'): ").strip()
    if not file_path:
        file_path = "inputs/data.xlsx"
    
    # Initialize reader
    reader = ExcelMultiSheetReader(file_path)
    
    # Load file
    if not reader.load_file():
        return
    
    # Main interactive loop
    while True:
        print("\n" + "-"*50)
        print("OPTIONS:")
        print("1. Read all sheets")
        print("2. Read specific sheet")
        print("3. Show sheets summary")
        print("4. Get definitions mapping")
        print("5. Get monitors for ID")
        print("6. Get conditions for ID")
        print("7. Export all to JSON")
        print("8. Export sheet to CSV")
        print("9. Exit")
        
        choice = input("\nEnter your choice (1-9): ").strip()
        
        if choice == "1":
            reader.read_all_sheets()
            
        elif choice == "2":
            sheet_name = input("Enter sheet name: ").strip()
            df = reader.read_specific_sheet(sheet_name)
            if df is not None:
                print(f"\nSheet '{sheet_name}' data:")
                print(df)
                
        elif choice == "3":
            reader.display_sheet_summary()
            
        elif choice == "4":
            mapping = reader.get_definitions_mapping()
            
        elif choice == "5":
            target_id = input("Enter ID to search for: ").strip()
            monitors_df = reader.get_monitors_for_id(target_id)
            if monitors_df is not None and not monitors_df.empty:
                print(f"\nMonitors for '{target_id}':")
                print(monitors_df)
                
        elif choice == "6":
            target_id = input("Enter ID to search for: ").strip()
            conditions_df = reader.get_conditions_for_id(target_id)
            if conditions_df is not None and not conditions_df.empty:
                print(f"\nConditions for '{target_id}':")
                print(conditions_df)
                
        elif choice == "7":
            output_file = input("Enter output JSON file path (default: outputs/excel_data.json): ").strip()
            if not output_file:
                output_file = "outputs/excel_data.json"
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            reader.export_to_json(output_file)
            
        elif choice == "8":
            sheet_name = input("Enter sheet name to export: ").strip()
            output_file = input(f"Enter output CSV file path (default: outputs/{sheet_name}.csv): ").strip()
            if not output_file:
                output_file = f"outputs/{sheet_name}.csv"
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            reader.export_sheet_to_csv(sheet_name, output_file)
            
        elif choice == "9":
            log_message("Goodbye!")
            break
            
        else:
            print("Invalid choice. Please try again.")


def demo_data_xlsx():
    """Demo function specifically for data.xlsx"""
    log_message("=== DEMO: Reading data.xlsx ===")
    
    reader = ExcelMultiSheetReader("inputs/data.xlsx")
    
    if not reader.load_file():
        return
    
    # Read all sheets
    reader.read_all_sheets()
    
    # Show summary
    reader.display_sheet_summary()
    
    # Get definitions mapping
    log_message("\n" + "="*40)
    log_message("DEFINITIONS MAPPING:")
    log_message("="*40)
    definitions = reader.get_definitions_mapping()
    
    # Show monitors and conditions for each ID
    if definitions:
        for fdir, id_val in definitions.items():
            log_message(f"\n--- Data for {fdir} (ID: {id_val}) ---")
            
            # Get monitors
            monitors = reader.get_monitors_for_id(id_val)
            if monitors is not None and not monitors.empty:
                log_message("Monitors:")
                for _, row in monitors.iterrows():
                    log_message(f"  {row['mons']}: {row['thresholds']}")
            
            # Get conditions
            conditions = reader.get_conditions_for_id(id_val)
            if conditions is not None and not conditions.empty:
                log_message("Conditions:")
                for _, row in conditions.iterrows():
                    log_message(f"  {row['condition_mons']}: count={row['counts']}, response={row['response']}")
    
    # Export to JSON
    os.makedirs("outputs", exist_ok=True)
    reader.export_to_json("outputs/data_xlsx_export.json")


def main():
    """Main entry point"""
    print("\nExcel Multi-Sheet Reader")
    print("Choose mode:")
    print("1. Interactive mode (explore any Excel file)")
    print("2. Demo mode (specifically for data.xlsx)")
    
    choice = input("\nEnter your choice (1/2): ").strip()
    
    if choice == "1":
        interactive_mode()
    elif choice == "2":
        demo_data_xlsx()
    else:
        print("Invalid choice.")


if __name__ == "__main__":
    main()