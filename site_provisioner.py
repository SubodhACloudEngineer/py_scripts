#!/usr/bin/env python3
"""
Generic Mist Site Provisioning Orchestrator
===========================================
Complete automation workflow: Excel → CSV → Mist Site Creation

This script orchestrates the entire site provisioning process:
1. Extracts site data from Excel files
2. Generates Mist-compatible CSV files
3. Creates sites in Mist organization with templates

Requirements:
- pandas, openpyxl (Excel processing)
- mistapi (Mist API)
- excel_to_mist_converter.py (Excel converter)
- mist_site_importer.py (Mist import script)
- config.json (site configuration)

Usage:
    python3 site_provisioner.py -f <excel_file> -s <site_id> [-o <org_id>]

Examples:
    # Interactive mode
    python3 site_provisioner.py -f site_data.xlsx -s SITE001
    
    # With specific org
    python3 site_provisioner.py -f site_data.xlsx -s SITE001 -o <org_id>
    
    # CSV only (no Mist import)
    python3 site_provisioner.py -f site_data.xlsx -s SITE001 --csv-only
"""

import sys
import os
import argparse
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def check_dependencies():
    """Check if required packages are installed"""
    required_packages = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl', 
        'mistapi': 'mistapi'
    }
    
    missing = []
    for package, import_name in required_packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(package)
    
    if missing:
        print("❌ Missing required packages:")
        for pkg in missing:
            print(f"  - {pkg}")
        print("\nInstall them with:")
        print(f"  pip install {' '.join(missing)}")
        sys.exit(1)

def check_required_files():
    """Check if required script files exist"""
    required_files = [
        'excel_to_mist_converter.py',
        'mist_site_importer.py',
        'config.json'
    ]
    
    missing = []
    for file in required_files:
        if not os.path.exists(file):
            missing.append(file)
    
    if missing:
        print("❌ Missing required files:")
        for file in missing:
            print(f"  - {file}")
        print("\nNote: Rename your Mist import script to 'mist_site_importer.py'")
        sys.exit(1)

def extract_site_data(excel_file, site_id, site_group="Default_Group"):
    """
    Extract site data from Excel file for specific site ID
    
    Args:
        excel_file: Path to Excel file
        site_id: Site identifier to extract
        site_group: Site group name for Mist
    
    Returns:
        CSV filename or None if failed
    """
    print("\n" + "="*80)
    print(" STEP 1: Extracting Site Data from Excel ".center(80, "="))
    print("="*80 + "\n")
    
    try:
        # Import the converter module
        import excel_to_mist_converter as converter
        
        logger.info(f"Processing Excel file: {excel_file}")
        logger.info(f"Looking for Site ID: {site_id}")
        
        # Call the extraction function
        success, csv_filename = converter.create_mist_csv(
            excel_file=excel_file,
            site_id=site_id,
            site_group=site_group,
            template_sheet="Site Variables",  # Customize if different
            start_marker=None  # Set to variable name if needed
        )
        
        if success and csv_filename:
            logger.info(f"CSV file created: {csv_filename}")
            return csv_filename
        else:
            logger.error("Failed to extract site data")
            return None
            
    except Exception as e:
        logger.error(f"Error during extraction: {str(e)}", exc_info=True)
        return None

def import_to_mist(csv_file, org_id=None, org_name=None, google_api_key=None):
    """
    Import the CSV file to Mist
    
    Args:
        csv_file: Path to CSV file
        org_id: Mist organization ID (optional)
        org_name: Mist organization name (optional)
        google_api_key: Google API key for geocoding (optional)
    
    Returns:
        True if successful, False otherwise
    """
    print("\n" + "="*80)
    print(" STEP 2: Importing Site to Mist ".center(80, "="))
    print("="*80 + "\n")
    
    try:
        # Import the Mist script module
        import mist_site_importer
        import mistapi
        
        # Create API session
        logger.info("Initializing Mist API session...")
        apisession = mistapi.APISession()
        apisession.login()
        
        # Call the start function from the Mist script
        logger.info(f"Starting site import from: {csv_file}")
        mist_site_importer.start(
            apisession=apisession,
            file_path=csv_file,
            org_id=org_id,
            org_name=org_name,
            google_api_key=google_api_key
        )
        
        logger.info("Site successfully imported to Mist!")
        return True
        
    except Exception as e:
        logger.error(f"Error during Mist import: {str(e)}", exc_info=True)
        return False

def cleanup(csv_file, keep_csv=False):
    """
    Clean up temporary files
    
    Args:
        csv_file: CSV file to potentially delete
        keep_csv: If True, keep the CSV file
    """
    if not keep_csv and os.path.exists(csv_file):
        try:
            response = input(f"\nDelete temporary CSV file '{csv_file}'? (y/n): ")
            if response.lower() == 'y':
                os.remove(csv_file)
                logger.info(f"Cleaned up: {csv_file}")
            else:
                logger.info(f"Keeping CSV file: {csv_file}")
        except Exception as e:
            logger.warning(f"Could not delete CSV file: {e}")

def main():
    """Main orchestration function"""
    parser = argparse.ArgumentParser(
        description='Complete Site Provisioning: Excel to Mist',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (prompts for org selection)
  python3 site_provisioner.py -f site_data.xlsx -s SITE001
  
  # With specific org ID
  python3 site_provisioner.py -f site_data.xlsx -s SITE001 -o <org_id>
  
  # With org name (creates new org if doesn't exist)
  python3 site_provisioner.py -f site_data.xlsx -s SITE001 -n "My Org"
  
  # Keep CSV file after import
  python3 site_provisioner.py -f site_data.xlsx -s SITE001 --keep-csv
  
  # CSV only (no Mist import)
  python3 site_provisioner.py -f site_data.xlsx -s SITE001 --csv-only
        """
    )
    
    parser.add_argument('-f', '--file', required=True,
                        help='Path to the Excel file containing site data')
    parser.add_argument('-s', '--site-id', required=True,
                        help='Site ID to extract from Excel')
    parser.add_argument('-g', '--site-group', default='Default_Group',
                        help='Site group name (default: Default_Group)')
    parser.add_argument('-o', '--org-id', 
                        help='Mist Organization ID (optional)')
    parser.add_argument('-n', '--org-name',
                        help='Mist Organization Name (optional)')
    parser.add_argument('--google-api-key',
                        help='Google API key for geocoding (optional)')
    parser.add_argument('--keep-csv', action='store_true',
                        help='Keep the generated CSV file after import')
    parser.add_argument('--csv-only', action='store_true',
                        help='Only generate CSV, do not import to Mist')
    
    args = parser.parse_args()
    
    # Print banner
    print("\n" + "="*80)
    print(" MIST SITE PROVISIONING AUTOMATION ".center(80, "="))
    print("="*80)
    print(f"\n  Excel File:  {args.file}")
    print(f"  Site ID:     {args.site_id}")
    print(f"  Site Group:  {args.site_group}")
    if args.org_id:
        print(f"  Org ID:      {args.org_id}")
    if args.org_name:
        print(f"  Org Name:    {args.org_name}")
    print()
    
    # Check dependencies
    logger.info("Checking dependencies...")
    check_dependencies()
    check_required_files()
    
    # Validate Excel file exists
    if not os.path.exists(args.file):
        logger.error(f"Excel file not found: {args.file}")
        sys.exit(1)
    
    # Step 1: Extract data from Excel
    csv_file = extract_site_data(args.file, args.site_id, args.site_group)
    
    if not csv_file:
        logger.error("Failed to generate CSV file")
        sys.exit(1)
    
    print(f"\n✓ CSV file generated successfully: {csv_file}")
    
    # If csv-only mode, stop here
    if args.csv_only:
        print(f"\n✓ CSV-only mode: Site data extracted to {csv_file}")
        print("  You can manually import this CSV or run the script again without --csv-only")
        sys.exit(0)
    
    # Ask for confirmation before proceeding to Mist
    print("\n" + "-"*80)
    response = input("Proceed with Mist import? (y/n): ")
    if response.lower() != 'y':
        print("Aborting. CSV file has been created for manual import.")
        sys.exit(0)
    
    # Step 2: Import to Mist
    success = import_to_mist(
        csv_file=csv_file,
        org_id=args.org_id,
        org_name=args.org_name,
        google_api_key=args.google_api_key
    )
    
    if success:
        print("\n" + "="*80)
        print(" SUCCESS! ".center(80, "="))
        print("="*80)
        print(f"\n✓ Site successfully provisioned in Mist")
        print(f"✓ Data extracted from Excel for Site ID: {args.site_id}")
        print(f"✓ Site created with all configurations applied\n")
        
        # Cleanup
        cleanup(csv_file, args.keep_csv)
        
    else:
        print("\n" + "="*80)
        print(" PARTIAL SUCCESS ".center(80, "="))
        print("="*80)
        print(f"\n✓ CSV file created: {csv_file}")
        print(f"✗ Mist import failed - check logs for details")
        print(f"\nYou can manually import the CSV file or retry with:")
        print(f"  python3 mist_site_importer.py -f {csv_file}\n")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        sys.exit(1)
