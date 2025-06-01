# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:percent
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.17.1
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

# %%
# Import configuration
import json
from datetime import datetime
import os
import pandas as pd
import pickle

def load_config(config_path="config.json"):
    """Load configuration from a JSON file."""
    with open(config_path, 'r') as file:
        config = json.load(file)
    return config["sqlserver_name"], config["sqlserver_db"], config["sqlserver_ip"], config["sqlserver_port"], config["sqlserver_user"], config["sqlserver_pwd"], config['base_url'], config['username'], config['password'], config['output_directory'], config['source_workbook_filename']
    

# Test loading configuration
sqlserver_name, sqlserver_db, sqlserver_ip, sqlserver_port, sqlserver_user, sqlserver_pwd, base_url, username, password, output_directory, source_workbook_filename = load_config()
print("Configuration loaded successfully.")

# %%
import requests
import urllib3

# Disable warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def authenticate(base_url, username, password):
    """Authenticate and return a bearer token."""
    auth_url = f"{base_url}/api/token"
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    payload = {
        "grant_type": "password",
        "username": username,
        "password": password
    }
    response = requests.post(auth_url, headers=headers, data=payload, verify=False)
    response.raise_for_status()
    token = response.json().get("access_token")  # Adjust if necessary
    return token

# Test authentication
token = authenticate(base_url, username, password)
print("Authentication successful. Token obtained.")

# %%
import requests
import urllib3
import os
import pickle

# Disable insecure certificate warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def fetch_paginated_collection(base_url, token, endpoint, page_size=50, cache_name=None):
    """
    Fetch an entire paginated collection from a REST API endpoint.
    
    Args:
        base_url (str): The base URL of the API.
        token (str): Bearer token for authorization.
        endpoint (str): The specific endpoint, e.g., '/api/teachers'.
        page_size (int): Number of records per page.
        cache_name (str, optional): If provided, caches the result under this name.

    Returns:
        Tuple[List[Dict], Dict]: The full result set and pagination metadata.
    """
    full_url = f"{base_url}{endpoint}?PageSize={page_size}"
    headers = {"Authorization": f"Bearer {token}"}
    page_no = 1
    all_items = []

    while True:
        paginated_url = f"{full_url}&PageNo={page_no}"
        response = requests.get(paginated_url, headers=headers, verify=False)
        response.raise_for_status()

        data = response.json()
        result_set = data.get("ResultSet", [])
        all_items.extend(result_set)

        pagination_info = {
            "HasPageInfo": data.get("HasPageInfo", False),
            "NumResults": data.get("NumResults", 0),
            "FirstRec": data.get("FirstRec", 1),
            "LastRec": data.get("LastRec", page_size),
            "PageSize": data.get("PageSize", page_size),
            "PageNo": data.get("PageNo", page_no),
            "IsLastPage": data.get("IsLastPage", True),
            "LastPage": data.get("LastPage", 1),
            "Tag": data.get("Tag", None)
        }

        if pagination_info["IsLastPage"]:
            break

        page_no += 1

    # Optionally cache the data
    if cache_name:
        os.makedirs("cached-data", exist_ok=True)
        with open(os.path.join("cached-data", f"{cache_name}.pkl"), "wb") as f:
            pickle.dump(all_items, f)
        print(f"✅ Cached {len(all_items)} items to cached-data/{cache_name}.pkl")

    return all_items, pagination_info



# %%
all_teachers, teacher_info = fetch_paginated_collection(
    base_url, token, "/api/teachers", page_size=50, cache_name="all_teachers"
)
all_schools, school_info = fetch_paginated_collection(
    base_url, token, "/api/schools", page_size=100, cache_name="all_schools"
)

# %store all_teachers all_schools

# %%
def get_lookups(base_url, token, lookup="core"):
    """Fetch lookups from the core collection endpoint."""
    url = f"{base_url}/api/lookups/collection/{lookup}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    
    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()
    
    lookups = response.json()
    return lookups  # This will likely be a dict or list of dicts depending on API

# Retrieve and store the lookups
core_lookups = get_lookups(base_url, token, "core")
student_lookups = get_lookups(base_url, token, "student")
censusworkbook_lookups = get_lookups(base_url, token, "censusworkbook")

# %store core_lookups student_lookups censusworkbook_lookups

# Optionally print keys or preview
print("Available core lookup categories:", list(core_lookups.keys()) if isinstance(core_lookups, dict) else type(core_lookups))
print("Available student lookup categories:", list(student_lookups.keys()) if isinstance(student_lookups, dict) else type(student_lookups))
print("Available censusworkbook lookup categories:", list(censusworkbook_lookups.keys()) if isinstance(censusworkbook_lookups, dict) else type(census_lookups))

# %%
censusworkbook_lookups['schoolCodes']

# %%
