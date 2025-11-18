**Mist Sites Mapper – README**

A lightweight, tenant-agnostic Python tool for retrieving all site IDs and site names from a Juniper Mist organization.
This script is fully generic — no client-specific information is embedded.
All configuration is supplied through environment variables or command-line flags.

**Features**

Retrieves all sites from a Mist organization using the standard API:
GET /api/v1/organizations/{org_id}/sites


**Prerequisites**
Python 3.8+

A Mist API token with Read permissions:
Organization Read
Site Read

**To generate a token:**
Mist Dashboard → Settings → API Tokens → Create token

**Installation**
Clone or copy the script:
mist_sites_mapper.py
README.md



**********************************************************
# Automated Mist Site Provisioning from Excel
**********************************************************

A complete Python-based automation solution for provisioning Juniper Mist sites from Excel metadata files.

## Overview

This solution automates the end-to-end process of creating sites in Juniper Mist from Excel spreadsheets containing site metadata. It eliminates manual data entry, reduces errors, and enables rapid site deployment at scale.

### Key Features

- **Intelligent Data Extraction**: Automatically searches Excel files to find site data based on unique identifiers
- **Mist CSV Generation**: Creates properly formatted CSV files compatible with Mist import requirements
- **Template Application**: Automatically applies Gateway, Network, and RF templates during site creation
- **Geocoding**: Automatically determines timezone, coordinates, and country code from addresses
- **Variable Management**: Supports 100+ site-specific variables (VLANs, subnets, etc.)
- **Batch Processing**: Can process multiple sites sequentially
- **Configuration Persistence**: Applies standardized site settings from config file

## Architecture

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│   Excel     │────>│  CSV Generator   │────>│ Mist API    │
│  Metadata   │     │  (Python Script) │     │  Importer   │
└─────────────┘     └──────────────────┘     └─────────────┘
                             │
                             ↓
                    ┌─────────────────┐
                    │  Mist Cloud     │
                    │  (Sites Created)│
                    └─────────────────┘
```

## Components

### 1. `excel_to_mist_converter.py`
- Extracts site data from Excel files
- Searches across all sheets for site identifiers
- Maps data to Mist CSV format
- Includes site variables from template sheet

### 2. `mist_site_importer.py`
- Imports CSV files to Mist via API
- Handles geocoding (Google Maps API or open-source)
- Applies templates (Gateway, Network, RF)
- Sets site configurations and variables
- Adds sites to site groups

### 3. `site_provisioner.py`
- Master orchestration script
- Combines Excel extraction + Mist import
- Provides unified command-line interface
- Handles error recovery and cleanup

### 4. `config.json`
- Standardized site settings
- Applied to all created sites
- Includes security policies, auto-upgrade settings, etc.

## Requirements

### Python Packages
```bash
pip install pandas openpyxl mistapi timezonefinder geopy
```

### Files Needed
- `excel_to_mist_converter.py` - Excel to CSV converter
- `mist_site_importer.py` - Mist import script
- `site_provisioner.py` - Master orchestrator
- `config.json` - Site configuration template

### Excel File Structure

Your Excel file should contain:

1. **Data Sheets**: One or more sheets with site data in rows
   - Site ID (unique identifier)
   - Address
   - Location/City
   - Additional site-specific data

2. **Template Sheet** (optional): Sheet named "Site Variables"
   - Variable names in Column A
   - Default values in Column B
   - Used to populate site variables in Mist

## Usage

### Quick Start

```bash
# Generate CSV only (review before importing)
python3 site_provisioner.py -f site_data.xlsx -s SITE001 --csv-only

# Full automation (Excel → CSV → Mist)
python3 site_provisioner.py -f site_data.xlsx -s SITE001
```

### Command Line Options

```bash
python3 site_provisioner.py \
  -f <excel_file>         # Required: Path to Excel file
  -s <site_id>            # Required: Site identifier to extract
  -g <site_group>         # Optional: Site group name (default: Default_Group)
  -o <org_id>             # Optional: Mist organization ID
  -n <org_name>           # Optional: Mist organization name
  --google-api-key <key>  # Optional: Google Maps API key for geocoding
  --csv-only              # Generate CSV only, don't import to Mist
  --keep-csv              # Keep CSV file after successful import
```

### Examples

**Example 1: Interactive Mode**
```bash
python3 site_provisioner.py -f sites.xlsx -s SITE001
# Script will prompt for organization selection
```

**Example 2: Specific Organization**
```bash
python3 site_provisioner.py \
  -f sites.xlsx \
  -s SITE001 \
  -o 12345678-1234-1234-1234-123456789abc
```

**Example 3: Custom Site Group**
```bash
python3 site_provisioner.py \
  -f sites.xlsx \
  -s SITE001 \
  -g "Branch_Offices"
```

**Example 4: Batch Processing**
```bash
# Process multiple sites
for site_id in SITE001 SITE002 SITE003; do
  python3 site_provisioner.py -f sites.xlsx -s $site_id -o <org_id>
done
```

## CSV Format

Generated CSV files follow Mist import requirements:

```csv
#name,address,sitegroup_names,vars
"Site_Name_SITE001","123 Main St, City, State","Branch_Offices","var1:value1,var2:value2,..."
```

### CSV Fields

**Required:**
- `name` - Site name (auto-generated: Location_SiteID)
- `address` - Full site address for geocoding

**Optional:**
- `sitegroup_names` - Comma-separated list of site groups
- `vars` - Site variables as key:value pairs
- `networktemplate_id` - Network template ID
- `gatewaytemplate_id` - Gateway template ID
- `rftemplate_id` - RF template ID
- `country_code` - Country code (auto-detected if not specified)
- `timezone` - Timezone (auto-detected if not specified)

## Configuration

### config.json Structure

```json
{
    "site_settings": {
        "persist_config_on_device": true,
        "rogue": {
            "min_rssi": -80,
            "min_duration": 10,
            "enabled": true,
            "honeypot_enabled": true
        },
        "auto_upgrade": {
            "enabled": true,
            "version": "custom"
        }
    }
}
```

### Customization

Edit `excel_to_mist_converter.py` to customize:

**Line 40**: Identifier field name
```python
id_field_name="Site ID"  # Change to match your Excel structure
```

**Line 41**: Minimum identifier length
```python
min_identifier_length=4  # Adjust for your ID format
```

**Line 52**: Sheet name patterns to skip
```python
if any(keyword in sheet_name.lower() for keyword in ['template', 'variables', 'config']):
```

**Line 130**: Column mapping logic
```python
if idx == target_col_idx:
    site_data['site_id'] = str(value)
elif idx == target_col_idx + 1:
    site_data['address'] = str(value)
elif idx == target_col_idx + 2:
    site_data['location'] = str(value)
```

## Troubleshooting

### Common Issues

**1. Module Not Found Error**
```bash
# Install missing packages
pip install pandas openpyxl mistapi timezonefinder geopy
```

**2. Geocoding Failures**
```bash
# Option 1: Use Google Maps API (more reliable)
python3 site_provisioner.py -f sites.xlsx -s SITE001 --google-api-key YOUR_KEY

# Option 2: Ensure open-source packages installed
pip install timezonefinder geopy
```

**3. Site ID Not Found**
```bash
# Run in CSV-only mode to see available IDs
python3 site_provisioner.py -f sites.xlsx -s TEST --csv-only
# Script will show available IDs in the Excel file
```

**4. Template Not Applied**
- Verify template IDs are correct in Mist organization
- Check that templates exist before site creation
- Use Mist UI to verify template assignments after import

### Debug Mode

Enable detailed logging:
```python
# In site_provisioner.py, change line 23:
logging.basicConfig(level=logging.DEBUG, ...)
```

## Best Practices

1. **Test First**: Always use `--csv-only` mode first to validate data extraction
2. **Review CSV**: Check generated CSV file before importing to Mist
3. **Backup**: Keep copies of Excel files and generated CSVs
4. **Incremental**: Process sites in batches, not all at once
5. **Validation**: Verify first few sites in Mist UI before batch processing
6. **Templates**: Create and test templates in Mist before automation
7. **Dry Run**: Test with non-production organization first

## Performance

- **CSV Generation**: ~2-5 seconds per site
- **Mist Import**: ~10-30 seconds per site (depends on geocoding)
- **Batch Processing**: ~15-40 seconds per site total

## Security

- **API Credentials**: Stored in `~/.mist_env` file (not in scripts)
- **No Hardcoded Secrets**: All sensitive data externalized
- **HTTPS Only**: All Mist API calls use encrypted connections

## Support

For issues or questions:
1. Check troubleshooting section above
2. Review log files (script.log)
3. Verify Excel file structure matches expected format
4. Test with minimal example first

## License

This solution is provided as-is for automation of Juniper Mist site provisioning.

## Version History

- **v1.0** - Initial release with Excel extraction and Mist import
- **v1.1** - Added batch processing support
- **v1.2** - Enhanced error handling and logging

---

**Note**: Customize this solution for your specific Excel structure and requirements. The scripts are designed to be flexible and adaptable to different data formats.
