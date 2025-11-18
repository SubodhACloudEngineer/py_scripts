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
