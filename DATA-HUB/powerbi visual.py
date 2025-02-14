import msal
import requests
import json
import pandas as pd
import time
import config

# 🔹 Microsoft Fabric Credentials
CLIENT_ID = config.client_id
CLIENT_SECRET = config.client_secret
TENANT_ID = config.tenant_id
WORKSPACE_ID =

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://api.fabric.microsoft.com/.default"]
EXCEL_FILE = "tableau_metadata3.xlsx"

def get_access_token():
    """Authenticate and get an access token for Fabric API"""
    app = msal.ConfidentialClientApplication(CLIENT_ID, client_credential=CLIENT_SECRET, authority=AUTHORITY)
    result = app.acquire_token_for_client(scopes=SCOPE)

    print("🔹 Full Token Response:", json.dumps(result, indent=2))  # Debugging

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("❌ Authentication failed: " + json.dumps(result, indent=2))

def read_excel_metadata():
    """Read Tableau metadata including visuals from the Excel file"""
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        raise Exception(f"❌ Excel file '{EXCEL_FILE}' not found.")

    # Check required columns
    required_columns = {"Workbook Name", "Dashboard Name", "Dashboard ID", "Dashboard URL", "Visual Type", "Columns Used"}
    if not required_columns.issubset(df.columns):
        raise Exception(f"❌ Missing required columns in '{EXCEL_FILE}'. Expected: {required_columns}")

    # Convert to Fabric format
    columns = [
        {"name": "Workbook Name", "dataType": "String"},
        {"name": "Dashboard Name", "dataType": "String"},
        {"name": "Dashboard ID", "dataType": "String"},
        {"name": "Dashboard URL", "dataType": "String"},
        {"name": "Visual Type", "dataType": "String"},
        {"name": "Columns Used", "dataType": "String"}
    ]

    # Convert rows into Fabric table format
    rows = df.to_dict(orient="records")  

    return columns, rows

def create_dataset():
    """Create a Fabric dataset using extracted metadata"""
    token = get_access_token()
    url = f"https://api.fabric.microsoft.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    columns, _ = read_excel_metadata()

    dataset_payload = {
        "name": "Tableau_Metadata_Dataset",
        "tables": [
            {
                "name": "TableauDashboards",
                "columns": columns
            }
        ]
    }

    response = requests.post(url, headers=headers, json=dataset_payload)

    print(f"🔍 API Response: {response.status_code}")
    print("🔍 Full API Response:", response.json())  # Debugging

    if response.status_code == 201:
        dataset_id = response.json()["id"]
        print(f"✅ Dataset created successfully: {dataset_id}")
        return dataset_id
    else:
        print(f"❌ Failed to create dataset: {response.text}")
        return None

def add_data_to_dataset(dataset_id):
    """Insert Tableau metadata rows into the Fabric dataset"""
    token = get_access_token()
    url = f"https://api.fabric.microsoft.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{dataset_id}/tables/TableauDashboards/rows"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    _, rows = read_excel_metadata()
    payload = {"rows": rows}

    response = requests.post(url, headers=headers, json=payload)
    
    print(f"🔍 API Response: {response.status_code}")
    print("🔍 Full API Response:", response.json())  # Debugging

    if response.status_code == 200:
        print("✅ Data added to dataset successfully.")
    else:
        print(f"❌ Failed to insert data: {response.text}")

def create_report(dataset_id):
    """Create a Fabric report in the workspace"""
    token = get_access_token()
    url = f"https://api.fabric.microsoft.com/v1.0/myorg/groups/{WORKSPACE_ID}/reports"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    report_payload = {
        "name": "Tableau_Migrated_Report",
        "datasetId": dataset_id
    }

    response = requests.post(url, headers=headers, json=report_payload)
    
    print(f"🔍 API Response: {response.status_code}")
    print("🔍 Full API Response:", response.json())  # Debugging

    if response.status_code == 201:
        report_id = response.json()["id"]
        print(f"✅ Report created successfully: {report_id}")
        return report_id
    else:
        print("❌ Failed to create report:", response.text)
        return None

def add_visuals(report_id):
    """Add visuals dynamically to the Fabric report based on Tableau metadata"""
    token = get_access_token()
    url = f"https://api.fabric.microsoft.com/v1.0/myorg/reports/{report_id}/pages/default/visuals"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    _, visuals = read_excel_metadata()

    for visual in visuals:
        visual_payload = {
            "visualType": visual["Visual Type"],
            "size": {"width": 400, "height": 300},
            "position": {"x": 50, "y": 50},
            "title": visual["Dashboard Name"],
            "columns": visual["Columns Used"].split(", ")  # Convert comma-separated columns into list
        }

        response = requests.post(url, headers=headers, json=visual_payload)
        
        print(f"🔍 API Response: {response.status_code}")
        print("🔍 Full API Response:", response.json())  # Debugging

        if response.status_code == 201:
            print(f"✅ Added {visual['Visual Type']} visual for {visual['Dashboard Name']}.")
        else:
            print(f"❌ Failed to add visual: {response.text}")

if __name__ == "__main__":
    print("🚀 Starting Tableau to Fabric Migration...")

    dataset_id = create_dataset()
    if dataset_id:
        time.sleep(5)  # Wait for dataset to be created
        add_data_to_dataset(dataset_id)
        
        report_id = create_report(dataset_id)
        if report_id:
            time.sleep(5)  # Wait for report to be created
            add_visuals(report_id)

    print("🎉 Migration completed! Check your Fabric workspace.")

