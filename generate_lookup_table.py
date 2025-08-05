#!/usr/bin/env python3
"""
Response Lookup Table Generator

This script reads response_text.xlsx and creates clean lookup tables
in multiple formats. Run this script whenever you update the Excel file.

Usage: python generate_lookup_table.py
"""

import pandas as pd
import json
import os
from datetime import datetime

def log_message(message):
    """Print timestamped log message"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def create_clean_lookup_table():
    """
    Main function to create clean lookup table from response_text.xlsx
    """
    # Configuration
    input_file = "inputs/response_text.xlsx"
    output_dir = "outputs"
    
    log_message("=== Response Lookup Table Generator ===")
    
    # Check if input file exists
    if not os.path.exists(input_file):
        log_message(f"ERROR: Input file '{input_file}' not found!")
        return False
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # Read the Excel file
        log_message(f"Reading {input_file}...")
        df = pd.read_excel(input_file)
        
        log_message(f"Loaded {len(df)} rows with columns: {list(df.columns)}")
        
        # Filter out rows where response is NaN
        df_clean = df.dropna(subset=['response'])
        log_message(f"Found {len(df_clean)} rows with valid responses")
        
        # Create lookup table - get unique response mappings
        response_cols = ['response', 'response_text', 'recovery_text']
        
        # Check if all required columns exist
        missing_cols = [col for col in response_cols if col not in df.columns]
        if missing_cols:
            log_message(f"ERROR: Missing columns: {missing_cols}")
            return False
        
        lookup_df = df_clean[response_cols].drop_duplicates()
        log_message(f"Creating lookup for {len(lookup_df)} unique response mappings")
        
        # Create fully clean dictionary lookup
        clean_lookup_dict = {}
        
        for _, row in lookup_df.iterrows():
            response_key = row['response']
            response_entry = {}
            
            # Parse response_text into lines
            if pd.notna(row['response_text']) and row['response_text']:
                response_lines = str(row['response_text']).replace('\\n', '\n').split('\n')
                response_text_lines = [line.strip() for line in response_lines if line.strip()]
                
                if response_text_lines:
                    # Single line stays as string, multiple lines become array
                    if len(response_text_lines) == 1:
                        response_entry["response_text"] = response_text_lines[0]
                    else:
                        response_entry["response_text_lines"] = response_text_lines
            
            # Parse recovery_text into steps array
            if pd.notna(row['recovery_text']) and row['recovery_text']:
                steps = str(row['recovery_text']).replace('\\n', '\n').split('\n')
                recovery_steps = [step.strip() for step in steps if step.strip()]
                
                if recovery_steps:
                    response_entry["recovery_steps"] = recovery_steps
            
            clean_lookup_dict[response_key] = response_entry
        
        # Display the clean lookup
        log_message("\n" + "="*60)
        log_message("GENERATED CLEAN LOOKUP TABLE:")
        log_message("="*60)
        
        for response, details in clean_lookup_dict.items():
            print(f"\nResponse: '{response}'")
            
            if 'response_text' in details:
                print(f"  - response_text: '{details['response_text']}'")
            elif 'response_text_lines' in details:
                print(f"  - response_text_lines: {len(details['response_text_lines'])} lines")
                for i, line in enumerate(details['response_text_lines'], 1):
                    print(f"      {i}. {line}")
            
            if 'recovery_steps' in details:
                print(f"  - recovery_steps: {len(details['recovery_steps'])} steps")
                for i, step in enumerate(details['recovery_steps'], 1):
                    print(f"      {i}. {step}")
        
        # Save in multiple formats
        log_message("\n" + "="*60)
        log_message("SAVING LOOKUP TABLES...")
        log_message("="*60)
        
        # 1. Save as Excel file with formatted data
        try:
            excel_file = os.path.join(output_dir, "response_lookup_table_updated.xlsx")
            
            # Create a formatted DataFrame for Excel export
            formatted_rows = []
            for response, details in clean_lookup_dict.items():
                row = {'response': response}
                
                # Format response text
                if 'response_text' in details:
                    row['response_text'] = details['response_text']
                elif 'response_text_lines' in details:
                    # Join lines with proper line breaks for Excel
                    row['response_text'] = '\n'.join(details['response_text_lines'])
                else:
                    row['response_text'] = ''
                
                # Format recovery steps (no numbering for Excel)
                if 'recovery_steps' in details:
                    # Just join steps with line breaks for Excel
                    row['recovery_steps'] = '\n'.join(details['recovery_steps'])
                else:
                    row['recovery_steps'] = ''
                
                formatted_rows.append(row)
            
            # Create DataFrame with formatted data
            formatted_df = pd.DataFrame(formatted_rows)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                formatted_df.to_excel(writer, sheet_name='Response Lookup', index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Response Lookup']
                
                # Set column widths
                worksheet.column_dimensions['A'].width = 25  # response
                worksheet.column_dimensions['B'].width = 40  # response_text
                worksheet.column_dimensions['C'].width = 50  # recovery_steps
                
                # Enable text wrapping for all cells
                from openpyxl.styles import Alignment
                wrap_alignment = Alignment(wrap_text=True, vertical='top')
                
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = wrap_alignment
                
                # Set row heights to auto-fit content
                for row_num in range(2, len(formatted_rows) + 2):  # Start from row 2 (skip header)
                    worksheet.row_dimensions[row_num].height = None  # Auto height
            
            log_message(f"✓ Saved formatted Excel: {excel_file}")
        except Exception as e:
            log_message(f"✗ Failed to save Excel: {e}")
            # Fallback to basic Excel save
            try:
                excel_file_basic = os.path.join(output_dir, "response_lookup_table_basic.xlsx")
                lookup_df.to_excel(excel_file_basic, index=False)
                log_message(f"✓ Saved basic Excel: {excel_file_basic}")
            except:
                pass
        
        # 2. Save as CSV file with clean single-line formatting
        try:
            csv_file = os.path.join(output_dir, "response_lookup_table_updated.csv")
            
            # Create CSV-friendly format (single line per field)
            csv_rows = []
            for response, details in clean_lookup_dict.items():
                row = {'response': response}
                
                # Format response text for CSV (use | separator for multiple lines)
                if 'response_text' in details:
                    row['response_text'] = details['response_text']
                elif 'response_text_lines' in details:
                    row['response_text'] = ' | '.join(details['response_text_lines'])
                else:
                    row['response_text'] = ''
                
                # Format recovery steps for CSV (use | separator)
                if 'recovery_steps' in details:
                    row['recovery_steps'] = ' | '.join(details['recovery_steps'])
                else:
                    row['recovery_steps'] = ''
                
                csv_rows.append(row)
            
            # Create CSV DataFrame
            csv_df = pd.DataFrame(csv_rows)
            csv_df.to_csv(csv_file, index=False)
            log_message(f"✓ Saved clean CSV: {csv_file}")
        except Exception as e:
            log_message(f"✗ Failed to save CSV: {e}")
        
        # 3. Save clean JSON file
        try:
            json_file = os.path.join(output_dir, "response_lookup_clean.json")
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(clean_lookup_dict, f, indent=2, ensure_ascii=False)
            log_message(f"✓ Saved clean JSON: {json_file}")
        except Exception as e:
            log_message(f"✗ Failed to save clean JSON: {e}")
        
        # 4. Save detailed JSON file (same as clean, for compatibility)
        try:
            detailed_json_file = os.path.join(output_dir, "response_lookup_detailed.json")
            with open(detailed_json_file, 'w', encoding='utf-8') as f:
                json.dump(clean_lookup_dict, f, indent=2, ensure_ascii=False)
            log_message(f"✓ Saved detailed JSON: {detailed_json_file}")
        except Exception as e:
            log_message(f"✗ Failed to save detailed JSON: {e}")
        
        # 5. Save fully clean JSON (alternative name)
        try:
            fully_clean_json = os.path.join(output_dir, "response_lookup_fully_clean.json")
            with open(fully_clean_json, 'w', encoding='utf-8') as f:
                json.dump(clean_lookup_dict, f, indent=2, ensure_ascii=False)
            log_message(f"✓ Saved fully clean JSON: {fully_clean_json}")
        except Exception as e:
            log_message(f"✗ Failed to save fully clean JSON: {e}")
        
        # Summary
        log_message("\n" + "="*60)
        log_message("SUMMARY:")
        log_message("="*60)
        log_message(f"✓ Input file: {input_file}")
        log_message(f"✓ Total original rows: {len(df)}")
        log_message(f"✓ Rows with valid responses: {len(df_clean)}")
        log_message(f"✓ Unique response mappings: {len(clean_lookup_dict)}")
        log_message(f"✓ Output directory: {output_dir}")
        log_message("✓ Formats: Excel (formatted), CSV (formatted), JSON (multiple versions)")
        log_message("\nClean format features:")
        log_message("  - Single-line response_text as string")
        log_message("  - Multi-line response_text as 'response_text_lines' array")
        log_message("  - Recovery steps as 'recovery_steps' array")
        log_message("  - No redundant fields")
        log_message("  - Conditional inclusion (only when data exists)")
        log_message("\nExcel formatting features:")
        log_message("  - Multi-line response_text with proper line breaks")
        log_message("  - Recovery steps with line breaks (no numbering)")
        log_message("  - Text wrapping enabled for readability")
        log_message("  - Auto-sized columns and rows")
        log_message("\nCSV formatting features:")
        log_message("  - Single-line format for clean display")
        log_message("  - Multi-line content separated by ' | '")
        log_message("  - Recovery steps separated by ' | '")
        
        return True
        
    except Exception as e:
        log_message(f"ERROR: Failed to process file: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main entry point"""
    success = create_clean_lookup_table()
    
    if success:
        log_message("\n🎉 Lookup table generation completed successfully!")
        
    else:
        log_message("\n❌ Lookup table generation failed!")
        exit(1)

if __name__ == "__main__":
    main()