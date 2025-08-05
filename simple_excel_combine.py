#!/usr/bin/env python3
"""
Simple Excel Combiner - Just the Combined Flat View
"""

import pandas as pd

def create_simple_combined_excel():
    """Create simple combined Excel with just the flat view"""
    
    print("Reading data.xlsx...")
    
    # Read all sheets
    definitions = pd.read_excel('inputs/data.xlsx', sheet_name='definitions')
    monitors = pd.read_excel('inputs/data.xlsx', sheet_name='monitors')  
    conditions = pd.read_excel('inputs/data.xlsx', sheet_name='conditions')
    
    print(f"Found {len(definitions)} definitions, {len(monitors)} monitors, {len(conditions)} conditions")
    
    # Get ALL unique IDs from monitors and conditions (not just definitions)
    all_ids = set(monitors['id'].unique()) | set(conditions['id'].dropna().unique())
    print(f"All IDs found: {sorted(all_ids)}")
    
    # Create FDIR name mapping (handle ID mismatches)
    id_to_fdir = {}
    for _, def_row in definitions.iterrows():
        id_to_fdir[def_row['id']] = def_row['FDIRs']
    
    # Handle ID mismatches and missing entries
    if 'fpu_bat' in all_ids and 'fpu_batt' in id_to_fdir:
        id_to_fdir['fpu_bat'] = 'Battery'  # Map fpu_bat to Battery
        print("⚠️  Fixed ID mismatch: fpu_bat -> Battery")
    
    if 'fpu_tracker' in all_ids:
        id_to_fdir['fpu_tracker'] = 'Tracker'  # Add missing tracker
        print("⚠️  Added missing: fpu_tracker -> Tracker")
    
    # Create combined flat data for ALL IDs
    combined_rows = []
    
    for fdir_id in sorted(all_ids):
        fdir_name = id_to_fdir.get(fdir_id, f"Unknown_{fdir_id}")
        
        print(f"Processing {fdir_name} ({fdir_id})...")
        
        # Get monitors and conditions for this FDIR
        fdir_monitors = monitors[monitors['id'] == fdir_id]
        fdir_conditions = conditions[conditions['id'] == fdir_id].dropna(subset=['condition_mons'])
        
        print(f"  - {len(fdir_monitors)} monitors, {len(fdir_conditions)} conditions")
        
        # Create rows combining monitors and conditions
        max_items = max(len(fdir_monitors), len(fdir_conditions), 1)
        
        for i in range(max_items):
            row = {
                'FDIRs': fdir_name,
                'id': fdir_id,
                'mons': fdir_monitors.iloc[i]['mons'] if i < len(fdir_monitors) else '',
                'thresholds': fdir_monitors.iloc[i]['thresholds'] if i < len(fdir_monitors) else '',
                'condition_mons': fdir_conditions.iloc[i]['condition_mons'] if i < len(fdir_conditions) else '',
                'counts': fdir_conditions.iloc[i]['counts'] if i < len(fdir_conditions) else '',
                'response': fdir_conditions.iloc[i]['response'] if i < len(fdir_conditions) else ''
            }
            combined_rows.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(combined_rows)
    
    # Load response lookup table and merge
    print("Loading response lookup table...")
    try:
        response_lookup = pd.read_excel('outputs/response_lookup_table_updated.xlsx')
        print(f"Found {len(response_lookup)} response definitions")
        
        # Merge with response lookup data
        df_merged = df.merge(response_lookup, on='response', how='left')
        
        # Reorder columns to keep original format + response data
        column_order = ['FDIRs', 'id', 'mons', 'thresholds', 'condition_mons', 'counts', 'response', 'response_text', 'recovery_steps']
        df_merged = df_merged[column_order]
        
        print(f"✅ Merged response lookup data successfully")
        
    except Exception as e:
        print(f"⚠️  Could not load response lookup table: {e}")
        print("Continuing without response lookup data...")
        df_merged = df
    
    # Save to Excel with proper newline formatting
    import os
    from datetime import datetime
    
    # Create a new filename if the original is in use
    base_output = 'outputs/combined_flat_with_responses.xlsx'
    output_file = base_output
    
    # If file exists and is locked, create a new version
    try:
        # Try to open for writing to check if locked
        with open(output_file, 'a'):
            pass
    except PermissionError:
        # File is locked, create new version with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f'outputs/combined_flat_with_responses_{timestamp}.xlsx'
        print(f"⚠️  Original file is open, creating: {output_file}")
    
    # Use ExcelWriter to control formatting
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_merged.to_excel(writer, sheet_name='Combined_Data', index=False)
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Combined_Data']
        
        # Import openpyxl formatting
        from openpyxl.styles import Alignment
        
        # Set text wrapping and alignment for all cells
        wrap_alignment = Alignment(wrap_text=True, vertical='top')
        
        # Apply formatting to all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = wrap_alignment
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        # Handle newlines in length calculation
                        cell_length = max(len(str(line)) for line in str(cell.value).split('\n'))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            
            # Set column width (cap at reasonable maximum)
            adjusted_width = min(max_length + 2, 60)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Set specific widths for known columns
        if 'H' in [cell.column_letter for cell in worksheet[1]]:  # response_text column
            worksheet.column_dimensions['H'].width = 40
        if 'I' in [cell.column_letter for cell in worksheet[1]]:  # recovery_steps column  
            worksheet.column_dimensions['I'].width = 50
        
        print(f"✅ Applied Excel formatting for proper newline display")
    
    print(f"✅ Created: {output_file}")
    print(f"📊 {len(df_merged)} rows with original columns + response lookup data")
    
    return df_merged

def main():
    """Main function"""
    print("📊 Simple Excel Combiner")
    print("=" * 30)
    
    df = create_simple_combined_excel()
    
    print("\nPreview:")
    print(df.to_string())

if __name__ == "__main__":
    main()