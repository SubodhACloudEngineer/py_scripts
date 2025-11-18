#!/usr/bin/env python3
"""
Generic Excel to Mist CSV Converter
====================================
Extracts site data from Excel files and converts to Mist-compatible CSV format.

This script intelligently searches through Excel sheets to find site data
based on a unique site identifier and generates properly formatted CSV files
for Mist site import.

Usage:
    python3 excel_to_mist_converter.py <excel_file>
    
The script will:
1. Scan all sheets in the Excel file for site data
2. Prompt for a site identifier (e.g., Site ID, Location Code)
3. Extract site information including location, address, and variables
4. Generate a Mist-compatible CSV file

Requirements:
    - pandas
    - openpyxl
"""

import pandas as pd
import sys
import os

def find_site_lookup_data(excel_file, id_field_name="Site ID", min_identifier_length=4):
    """
    Find the lookup table that contains all site identifiers and their corresponding data
    
    Args:
        excel_file: Path to Excel file
        id_field_name: Name of the identifier field to search for (default: "Site ID")
        min_identifier_length: Minimum length for valid identifiers (default: 4)
    
    Returns:
        Dictionary containing lookup data from all sheets
    """
    print(f"🔍 Searching for site data in Excel file...")
    print(f"   Looking for identifier field: '{id_field_name}'")
    
    xl_file = pd.ExcelFile(excel_file)
    sheet_names = xl_file.sheet_names
    print(f"   Found sheets: {sheet_names}\n")
    
    lookup_data = {}
    
    for sheet_name in sheet_names:
        print(f"--- Examining '{sheet_name}' sheet ---")
        
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            print(f"Sheet size: {len(df)} rows x {len(df.columns)} columns")
            
            # Skip template sheets (usually contain "Template" or "Variables" in name)
            if any(keyword in sheet_name.lower() for keyword in ['template', 'variables', 'config']):
                print(f"Skipping template/config sheet: {sheet_name}")
                continue
            
            # Look for site identifier patterns
            identifiers_found = []
            potential_rows = []
            
            for idx, row in df.iterrows():
                for col_idx, value in enumerate(row):
                    # Look for identifiers (numeric or alphanumeric with min length)
                    if pd.notna(value):
                        value_str = str(value).strip()
                        if len(value_str) >= min_identifier_length:
                            identifiers_found.append(value_str)
                            potential_rows.append((idx, col_idx, value_str))
            
            if len(identifiers_found) > 2:
                unique_ids = list(set(identifiers_found))
                print(f"Found {len(unique_ids)} potential site identifiers")
                print(f"Sample IDs: {unique_ids[:5]}")
                
                lookup_data[sheet_name] = {
                    'dataframe': df,
                    'identifiers': unique_ids,
                    'potential_rows': potential_rows
                }
        
        except Exception as e:
            print(f"Could not read sheet '{sheet_name}': {e}")
    
    return lookup_data

def extract_site_data_from_lookup(lookup_data, target_id):
    """
    Extract site data for a specific identifier from the lookup tables
    
    Args:
        lookup_data: Dictionary containing sheet data
        target_id: Target site identifier to find
    
    Returns:
        Tuple of (site_data dict, source_sheet name) or (None, None) if not found
    """
    print(f"\n🎯 Looking for site identifier: {target_id}")
    
    for sheet_name, sheet_data in lookup_data.items():
        print(f"\nChecking '{sheet_name}' sheet...")
        
        if target_id in sheet_data['identifiers']:
            print(f"✓ Found identifier {target_id} in '{sheet_name}'!")
            
            df = sheet_data['dataframe']
            
            # Find the row and column with this identifier
            target_row_idx = None
            target_col_idx = None
            
            for row_idx, col_idx, identifier in sheet_data['potential_rows']:
                if identifier == target_id:
                    target_row_idx = row_idx
                    target_col_idx = col_idx
                    break
            
            if target_row_idx is not None:
                print(f"Found at row {target_row_idx}, column {target_col_idx}")
                
                row_data = df.iloc[target_row_idx]
                site_data = {}
                
                # Extract adjacent columns (common pattern: ID, Address, Location, ...)
                for idx, value in enumerate(row_data):
                    if pd.notna(value) and value != '':
                        if isinstance(value, str) and len(value) > 2:
                            # Map based on typical column patterns
                            if idx == target_col_idx:
                                site_data['site_id'] = str(value)
                            elif idx == target_col_idx + 1:
                                site_data['address'] = str(value)
                            elif idx == target_col_idx + 2:
                                site_data['location'] = str(value)
                
                print(f"Extracted data: {site_data}")
                return site_data, sheet_name
    
    return None, None

def get_site_variables_template(excel_file, template_sheet="Site Variables", start_marker=None):
    """
    Get site variable templates from a designated template sheet
    
    Args:
        excel_file: Path to Excel file
        template_sheet: Name of the sheet containing variable templates
        start_marker: Variable name to start collecting from (optional)
    
    Returns:
        List of (variable_name, default_value) tuples
    """
    try:
        print(f"📋 Getting variable template from '{template_sheet}' sheet...")
        df_template = pd.read_excel(excel_file, sheet_name=template_sheet, header=None)
        
        template_vars = []
        vars_started = False if start_marker else True
        
        for idx, row in df_template.iterrows():
            if pd.notna(row[0]) and isinstance(row[0], str):
                var_name = row[0].strip()
                
                # Skip header rows
                if any(keyword in var_name.lower() for keyword in ['branch type', 'template', 'type:']):
                    continue
                
                # Check if we should start collecting
                if start_marker and var_name == start_marker:
                    vars_started = True
                
                if vars_started:
                    template_value = row[1] if pd.notna(row[1]) else '0'
                    if template_value not in ['', 'undefined']:
                        template_vars.append((var_name, str(template_value)))
        
        print(f"Loaded {len(template_vars)} variable templates")
        return template_vars
        
    except Exception as e:
        print(f"Warning: Could not load variable template: {e}")
        return []

def create_mist_csv(excel_file, site_id, output_csv=None, 
                   site_group="Default_Group", template_sheet="Site Variables",
                   start_marker=None):
    """
    Create Mist-compatible CSV by finding the actual data for the specified site ID
    
    Args:
        excel_file: Path to Excel file
        site_id: Site identifier to extract
        output_csv: Output CSV filename (optional, auto-generated if None)
        site_group: Site group name for Mist
        template_sheet: Name of sheet containing variable templates
        start_marker: Variable name to start including in vars field
    
    Returns:
        Tuple of (success: bool, csv_filename: str or None)
    """
    try:
        # Find lookup data
        lookup_data = find_site_lookup_data(excel_file)
        
        if not lookup_data:
            print("❌ No site data found in Excel file")
            return False, None
        
        # Extract site-specific data
        site_data, source_sheet = extract_site_data_from_lookup(lookup_data, site_id)
        
        if not site_data:
            print(f"❌ Could not find data for site ID: {site_id}")
            available_ids = []
            for sheet_data in lookup_data.values():
                available_ids.extend(sheet_data['identifiers'][:10])
            print(f"Sample available IDs: {available_ids[:20]}")
            return False, None
        
        print(f"\n✅ Found site data in '{source_sheet}' sheet")
        
        # Get variable templates
        template_vars = get_site_variables_template(excel_file, template_sheet, start_marker)
        
        # Build Mist site entry
        location = site_data.get('location', f'Site_{site_id}')
        address = site_data.get('address', 'Address not specified')
        
        mist_site = {
            'name': f"{location}_{site_id}",
            'address': address,
            'sitegroup_names': site_group
        }
        
        # Build vars string
        if template_vars:
            vars_list = []
            for var_name, var_value in template_vars:
                clean_var_name = var_name.replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '')
                vars_list.append(f"{clean_var_name}:{var_value}")
            
            mist_site['vars'] = '"' + ','.join(vars_list) + '"'
        
        # Create output file
        if output_csv is None:
            clean_location = location.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '_')
            output_csv = f"{clean_location}_{site_id}.csv"
        
        # Write CSV
        output_df = pd.DataFrame([mist_site])
        
        with open(output_csv, 'w', encoding='utf-8') as f:
            header_line = '#' + ','.join(output_df.columns)
            f.write(header_line + '\n')
            output_df.to_csv(f, index=False, header=False, quoting=1)
        
        print(f"\n✅ Successfully created: {output_csv}")
        print(f"📍 Site: {location} (ID: {site_id})")
        if template_vars:
            print(f"📋 Variables: {len(template_vars)} from template")
        
        return True, output_csv
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False, None

def main():
    """Main function for standalone use"""
    print("\n" + "="*70)
    print(" Generic Excel to Mist CSV Converter ".center(70, "="))
    print("="*70 + "\n")
    
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python3 excel_to_mist_converter.py <excel_file>")
        print()
        print("Example:")
        print("  python3 excel_to_mist_converter.py site_data.xlsx")
        print()
        return
    
    excel_file = sys.argv[1]
    
    if not os.path.exists(excel_file):
        print(f"❌ Error: File '{excel_file}' not found.")
        return
    
    # Configuration (customize these for your Excel structure)
    print("Configuration:")
    print("  Default template sheet: 'Site Variables'")
    print("  Default site group: 'Default_Group'")
    print()
    
    # Get site ID from user
    while True:
        site_id = input("💬 Enter the Site ID: ").strip()
        if site_id and len(site_id) >= 4:
            break
        else:
            print("❌ Please enter a valid Site ID (4+ characters)")
    
    print(f"\n🚀 Processing Site ID: {site_id}\n")
    
    # Call the main extraction function
    success, csv_file = create_mist_csv(
        excel_file=excel_file,
        site_id=site_id,
        site_group="Default_Group",
        template_sheet="Site Variables",
        start_marker=None  # Set to a variable name if you want to skip initial vars
    )
    
    if success and csv_file:
        print(f"\n🎉 Conversion completed successfully!")
        print(f"📄 CSV file: {csv_file}")
        print(f"\nYou can now import this CSV to Mist using:")
        print(f"  - Mist Web UI: Organization > Sites > Import")
        print(f"  - Or use the site import script")
    else:
        print(f"\n💥 Conversion failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
