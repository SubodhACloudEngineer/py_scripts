#!/usr/bin/env python3
"""
Generic Mist Sites Mapper
- Lists all sites for a Mist organization and returns {site_id: site_name}.
- No client-specific IDs/URLs/tokens are hardcoded.
- Works across regions by setting MIST_BASE_URL.

Inputs (env or CLI):
  MIST_API_TOKEN   (required)  -> e.g., export MIST_API_TOKEN="xxxx"
  MIST_ORG_ID      (optional)  -> if omitted, script uses the first org you can access via /self/orgs
  MIST_BASE_URL    (optional)  -> default: https://api.mist.com/api/v1
                                   EU example: https://api.eu.mist.com/api/v1

Examples:
  python3 mist_sites_mapper.py --format json
  python3 mist_sites_mapper.py --format csv --outfile sites.csv
  MIST_BASE_URL="https://api.eu.mist.com/api/v1" python3 mist_sites_mapper.py
"""

import os
import sys
import json
import csv
import time
import argparse
from typing import Dict, List, Any, Tuple
import requests

DEFAULT_BASE_URL = "https://api.mist.com/api/v1"
PAGE_LIMIT = 100
MAX_RETRIES = 5
RETRY_BASE_S = 1.5

def build_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Token {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    })
    return s

def req_with_backoff(s: requests.Session, method: str, url: str, **kwargs) -> Any:
    for attempt in range(1, MAX_RETRIES + 1):
        r = s.request(method, url, timeout=60, **kwargs)
        if r.status_code in (200, 201):
            try:
                return r.json()
            except Exception:
                return {}
        if r.status_code in (429, 500, 502, 503, 504):
            sleep_s = min(30.0, RETRY_BASE_S * (2 ** (attempt - 1)))
            time.sleep(sleep_s)
            continue
        # Other errors -> raise with context (no secrets)
        raise requests.HTTPError(f"{r.status_code} on {url} -> {r.text[:300]}", response=r)
    raise RuntimeError(f"Gave up after {MAX_RETRIES} retries on {url}")

def choose_org_id(s: requests.Session, base_url: str, provided_org_id: str = None) -> str:
    if provided_org_id:
        return provided_org_id
    js = req_with_backoff(s, "GET", f"{base_url}/self/orgs")
    if not isinstance(js, list) or not js:
        raise RuntimeError("No accessible orgs for this token; provide MIST_ORG_ID.")
    return js[0]["id"]

def list_sites(s: requests.Session, base_url: str, org_id: str) -> List[Dict[str, Any]]:
    """Standard, tenant-agnostic listing with pagination support."""
    sites: List[Dict[str, Any]] = []
    page = 1
    while True:
        url = f"{base_url}/organizations/{org_id}/sites?limit={PAGE_LIMIT}&page={page}"
        js = req_with_backoff(s, "GET", url)
        items = js.get("results", js) if isinstance(js, dict) else js
        if not isinstance(items, list) or not items:
            break
        sites.extend(items)
        # stop if not paginated or last page
        if not (isinstance(js, dict) and "results" in js and len(items) == PAGE_LIMIT):
            break
        page += 1
    return sites

def to_mapping(sites: List[Dict[str, Any]]) -> Dict[str, str]:
    return {site.get("id"): site.get("name", "") for site in sites if site.get("id")}

def write_csv(mapping: Dict[str, str], outfile: str) -> None:
    with open(outfile, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["site_id", "site_name"])
        for sid, name in mapping.items():
            w.writerow([sid, name])

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generic Mist site ID→name mapper")
    p.add_argument("--base-url", default=os.getenv("MIST_BASE_URL", DEFAULT_BASE_URL),
                   help=f"Mist API base URL (default: {DEFAULT_BASE_URL})")
    p.add_argument("--org-id", default=os.getenv("MIST_ORG_ID"),
                   help="Organization ID (if omitted, uses first accessible org)")
    p.add_argument("--token", default=os.getenv("MIST_API_TOKEN"),
                   help="Mist API token (or set MIST_API_TOKEN env var)")
    p.add_argument("--format", choices=["json", "csv"], default="json",
                   help="Output format (default: json)")
    p.add_argument("--outfile", help="Path to CSV file (required if --format csv)")
    return p.parse_args()

def main():
    args = parse_args()

    if not args.token:
        print("ERROR: Provide Mist API token via --token or MIST_API_TOKEN env var.", file=sys.stderr)
        sys.exit(2)

    s = build_session(args.token)

    try:
        org_id = choose_org_id(s, args.base_url, args.org_id)
        sites = list_sites(s, args.base_url, org_id)
        mapping = to_mapping(sites)
    except requests.HTTPError as e:
        # Keep error readable but secret-safe
        print(f"HTTP error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if args.format == "json":
        print(json.dumps(mapping, ensure_ascii=False, indent=2))
    else:
        if not args.outfile:
            print("ERROR: --outfile is required for CSV output.", file=sys.stderr)
            sys.exit(2)
        write_csv(mapping, args.outfile)
        print(f"Wrote CSV: {args.outfile}")

if __name__ == "__main__":
    main()
