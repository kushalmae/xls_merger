#!/usr/bin/env python3
"""
Simple Data Combiner

Combines data from all sheets in data.xlsx to show a holistic view
of each FDIR with its monitors and conditions.
"""

import pandas as pd
import json

def combine_excel_data():
    """Simple function to combine all sheet data"""
    
    # Read all sheets
    print("Reading data.xlsx...")
    definitions = pd.read_excel('inputs/data.xlsx', sheet_name='definitions')
    monitors = pd.read_excel('inputs/data.xlsx', sheet_name='monitors')  
    conditions = pd.read_excel('inputs/data.xlsx', sheet_name='conditions')
    
    print(f"✓ Definitions: {len(definitions)} rows")
    print(f"✓ Monitors: {len(monitors)} rows") 
    print(f"✓ Conditions: {len(conditions)} rows")
    
    # Create combined data structure
    combined_data = {}
    
    # Process each FDIR
    for _, row in definitions.iterrows():
        fdir_name = row['FDIRs']
        fdir_id = row['id']
        
        print(f"\nProcessing {fdir_name} (ID: {fdir_id})...")
        
        # Get monitors for this ID
        fdir_monitors = monitors[monitors['id'] == fdir_id]
        monitor_list = []
        for _, mon_row in fdir_monitors.iterrows():
            monitor_list.append({
                'monitor': mon_row['mons'],
                'threshold': mon_row['thresholds']
            })
        
        # Get conditions for this ID (remove NaN rows)
        fdir_conditions = conditions[conditions['id'] == fdir_id].dropna(subset=['condition_mons'])
        condition_list = []
        for _, cond_row in fdir_conditions.iterrows():
            condition_list.append({
                'condition': cond_row['condition_mons'],
                'count': cond_row['counts'],
                'response': cond_row['response']
            })
        
        # Combine everything for this FDIR
        combined_data[fdir_name] = {
            'id': fdir_id,
            'monitors': monitor_list,
            'conditions': condition_list
        }
        
        print(f"  - {len(monitor_list)} monitors")
        print(f"  - {len(condition_list)} conditions")
    
    return combined_data

def display_combined_data(data):
    """Display the combined data in a readable format"""
    
    print("\n" + "="*60)
    print("HOLISTIC VIEW - ALL FDIRS")
    print("="*60)
    
    for fdir_name, fdir_data in data.items():
        print(f"\n📡 {fdir_name.upper()} (ID: {fdir_data['id']})")
        print("-" * 50)
        
        # Show monitors
        print("🔍 MONITORS:")
        if fdir_data['monitors']:
            for mon in fdir_data['monitors']:
                print(f"  • {mon['monitor']}: {mon['threshold']}")
        else:
            print("  (No monitors)")
        
        # Show conditions  
        print("\n⚡ CONDITIONS:")
        if fdir_data['conditions']:
            for cond in fdir_data['conditions']:
                print(f"  • {cond['condition']}")
                print(f"    Count: {cond['count']}, Response: {cond['response']}")
        else:
            print("  (No conditions)")

def save_combined_data(data):
    """Save combined data to files"""
    
    print("\n" + "="*40)
    print("SAVING COMBINED DATA")
    print("="*40)
    
    # Save as JSON
    with open('outputs/combined_data.json', 'w') as f:
        json.dump(data, f, indent=2)
    print("✓ Saved: outputs/combined_data.json")
    
    # Save as readable text file
    with open('outputs/combined_data.txt', 'w') as f:
        f.write("HOLISTIC VIEW - ALL FDIRS\n")
        f.write("="*60 + "\n")
        
        for fdir_name, fdir_data in data.items():
            f.write(f"\n{fdir_name.upper()} (ID: {fdir_data['id']})\n")
            f.write("-" * 50 + "\n")
            
            f.write("MONITORS:\n")
            if fdir_data['monitors']:
                for mon in fdir_data['monitors']:
                    f.write(f"  • {mon['monitor']}: {mon['threshold']}\n")
            else:
                f.write("  (No monitors)\n")
            
            f.write("\nCONDITIONS:\n")
            if fdir_data['conditions']:
                for cond in fdir_data['conditions']:
                    f.write(f"  • {cond['condition']}\n")
                    f.write(f"    Count: {cond['count']}, Response: {cond['response']}\n")
            else:
                f.write("  (No conditions)\n")
    
    print("✓ Saved: outputs/combined_data.txt")
    
    # Save as CSV (flattened)
    csv_rows = []
    for fdir_name, fdir_data in data.items():
        # Create a row for each monitor/condition combination
        if fdir_data['monitors'] or fdir_data['conditions']:
            max_items = max(len(fdir_data['monitors']), len(fdir_data['conditions']))
            
            for i in range(max_items):
                row = {
                    'FDIR': fdir_name,
                    'ID': fdir_data['id'],
                    'Monitor': fdir_data['monitors'][i]['monitor'] if i < len(fdir_data['monitors']) else '',
                    'Threshold': fdir_data['monitors'][i]['threshold'] if i < len(fdir_data['monitors']) else '',
                    'Condition': fdir_data['conditions'][i]['condition'] if i < len(fdir_data['conditions']) else '',
                    'Count': fdir_data['conditions'][i]['count'] if i < len(fdir_data['conditions']) else '',
                    'Response': fdir_data['conditions'][i]['response'] if i < len(fdir_data['conditions']) else ''
                }
                csv_rows.append(row)
        else:
            # Empty row if no monitors or conditions
            csv_rows.append({
                'FDIR': fdir_name,
                'ID': fdir_data['id'],
                'Monitor': '', 'Threshold': '', 'Condition': '', 'Count': '', 'Response': ''
            })
    
    csv_df = pd.DataFrame(csv_rows)
    csv_df.to_csv('outputs/combined_data.csv', index=False)
    print("✓ Saved: outputs/combined_data.csv")

def main():
    """Main function"""
    print("🔄 Data Combiner - Creating Holistic View")
    print("="*50)
    
    # Combine the data
    combined = combine_excel_data()
    
    # Display it
    display_combined_data(combined)
    
    # Save it
    save_combined_data(combined)
    
    print(f"\n🎉 Successfully combined data for {len(combined)} FDIRs!")
    print("Check the outputs folder for saved files.")

if __name__ == "__main__":
    main()