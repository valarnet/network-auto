import pandas as pd
import numpy as np
import os
import time

def main():
    # Create int_parsed_outputs directory if it doesn't exist
    output_dir = "int_parsed_outputs"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
    
    input_file = input('Enter the path to the input Excel file: ')
    
    # Generate output filename with timestamp
    input_filename = os.path.basename(input_file)
    base, ext = os.path.splitext(input_filename)
    
    # Extract the part before the first underscore
    # For example, from "stc-switches_show_int_status_parsed_20250517" get "stc-switches"
    parts = base.split('_')
    output_prefix = parts[0]
    
    current_date = time.strftime("%Y%m%d")
    
    # Generate unique filename with date and sequence number if needed
    output_base = os.path.join(output_dir, f"{output_prefix}_active_physical_intf_count_{current_date}")
    
    # Check if files with this date already exist and add sequence number if needed
    seq_num = 1
    output_file = f"{output_base}.xlsx"
    
    while os.path.exists(output_file):
        seq_num += 1
        output_file = f"{output_base}_{seq_num}.xlsx"
    
    print(f"Output will be saved to {output_file}")

    try:
        all_sheets = pd.read_excel(input_file, sheet_name=None)
    except Exception as e:
        print(f"Error reading input file: {e}")
        return

    # Prepare lists to hold each sheet's summary DataFrames
    speed_summaries = []
    type_summaries = []

    # Keep track of processed sheets to add blank rows between them
    processed_sheets = 0

    for sheet_name, df in all_sheets.items():
        if 'Status' in df.columns and 'Speed' in df.columns and 'Name' in df.columns:
            # Check if Type column exists
            has_type_column = 'Type' in df.columns
            
            # Filter for connected interfaces and create a copy to avoid SettingWithCopyWarning
            connected = df[df['Status'].astype(str).str.lower() == 'connected'].copy()
            
            # Convert names to lowercase for case-insensitive comparison
            connected.loc[:, 'Name_Lower'] = connected['Name'].astype(str).str.lower().str.strip()
            connected.loc[:, 'Speed_Lower'] = connected['Speed'].astype(str).str.lower().str.strip()
            
            # Create name filter mask
            name_filter = ~(connected['Name_Lower'].str.startswith('po') | 
                           connected['Name_Lower'].str.startswith('lo') | 
                           connected['Name_Lower'].str.startswith('vlan') |
                           connected['Name_Lower'].str.startswith('nve'))
            
            # Calculate how many interfaces were excluded by name
            name_excluded_count = len(connected) - name_filter.sum()
            
            # Create speed filter mask
            speed_filter = ~connected['Speed_Lower'].str.contains('auto')
            
            # Apply both filters
            combined_filter = name_filter & speed_filter
            filtered_connected = connected[combined_filter]
            
            # Calculate how many interfaces were excluded by speed
            speed_excluded_count = name_filter.sum() - combined_filter.sum()
            
            # Count by Speed
            speed_counts = filtered_connected['Speed'].value_counts(dropna=False)
            speed_summary_df = speed_counts.reset_index().rename(
                columns={'index': 'Speed', 'Speed': 'Count'}
            )
            speed_summary_df.insert(0, 'Switch', sheet_name)
            
            # Add a blank row after this sheet's data (except for the first sheet)
            if processed_sheets > 0:
                # Create a blank row DataFrame with the same columns
                blank_row = pd.DataFrame([[np.nan, np.nan, np.nan]], 
                                        columns=['Switch', 'Speed', 'Count'])
                # Add the blank row before this sheet's data
                speed_summary_df = pd.concat([blank_row, speed_summary_df], ignore_index=True)
            
            speed_summaries.append(speed_summary_df)
            
            # Count by Type if the column exists
            if has_type_column:
                type_counts = filtered_connected['Type'].value_counts(dropna=False)
                type_summary_df = type_counts.reset_index().rename(
                    columns={'index': 'Type', 'Type': 'Count'}
                )
                type_summary_df.insert(0, 'Switch', sheet_name)
                
                # Add a blank row after this sheet's data (except for the first sheet)
                if processed_sheets > 0:
                    # Create a blank row DataFrame with the same columns
                    blank_row = pd.DataFrame([[np.nan, np.nan, np.nan]], 
                                            columns=['Switch', 'Type', 'Count'])
                    # Add the blank row before this sheet's data
                    type_summary_df = pd.concat([blank_row, type_summary_df], ignore_index=True)
                
                type_summaries.append(type_summary_df)
            
            processed_sheets += 1
        else:
            # Skip sheets without required columns silently
            pass

    # Check if we have any data to write
    if not speed_summaries and not type_summaries:
        print("No valid data found in the input file. Make sure it contains sheets with 'Status', 'Speed', and 'Name' columns.")
        return

    # Write all summaries to Excel sheets
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Combine all speed summaries into a single DataFrame and write to Excel
            if speed_summaries:
                combined_speed_df = pd.concat(speed_summaries, ignore_index=True)
                combined_speed_df.to_excel(writer, sheet_name='Speed Summary', index=False)
            
            # Combine all type summaries into a single DataFrame and write to Excel
            if type_summaries:
                combined_type_df = pd.concat(type_summaries, ignore_index=True)
                combined_type_df.to_excel(writer, sheet_name='Type Summary', index=False)
        
        # Add a success message with the output file path
        print(f"File processed successfully. Output written to: {output_file}")
    except Exception as e:
        print(f"Error writing output file: {e}")

if __name__ == "__main__":
    main()
