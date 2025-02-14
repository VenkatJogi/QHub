import requests
import pandas as pd
import config

# 🔹 Microsoft Fabric (Power BI) Credentials
CLIENT_ID = config.client_id
CLIENT_SECRET = config.client_secret
TENANT_ID = config.tenant_id
WORKSPACE_ID =
EXCEL_FILE = 
DATASET_NAME = 

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# 🔹 Get Access Token for Power BI API
def get_access_token():
    url = f"{AUTHORITY}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(url, data=data, headers=headers)
    token = response.json().get("access_token")
    if not token:
        raise Exception(f"❌ Failed to get access token: {response.text}")
    return token

# 🔹 Get Dataset ID Dynamically
def get_dataset_id():
    token = get_access_token()
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        datasets = response.json().get("value", [])
        for dataset in datasets:
            if dataset["name"] == DATASET_NAME:
                return dataset["id"]
    return None  # Dataset does not exist

# 🔹 Create Dataset if it Doesn't Exist
def create_dataset():
    token = get_access_token()
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    dataset_payload = {
        "name": DATASET_NAME,
        "defaultMode": "Push",
        "tables": [
            {
                "name": "TableauData",
                "columns": [
                    {"name": "WorkbookName", "dataType": "String"},
                    {"name": "DashboardName", "dataType": "String"},
                    {"name": "DashboardID", "dataType": "String"},
                    {"name": "DashboardURL", "dataType": "String"},
                    {"name": "VisualType", "dataType": "String"},
                    {"name": "ColumnsUsed", "dataType": "String"}
                ]
            }
        ]
    }

    response = requests.post(url, headers=headers, json=dataset_payload)

    if response.status_code == 201:
        print(f"✅ Dataset '{DATASET_NAME}' created successfully.")
        return get_dataset_id()
    else:
        print(f"❌ Failed to create dataset. Response Code: {response.status_code}")
        print(f"❌ Response Text: {response.text}")  # Debugging Output
        raise Exception("Dataset creation failed")

# 🔹 Upload Data to Dataset
def upload_data_to_power_bi():
    df = pd.read_excel(EXCEL_FILE)

    dataset_id = get_dataset_id() or create_dataset()  # Get dataset ID (or create it)
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{dataset_id}/tables/TableauData/rows"
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {get_access_token()}"
    }

    rows = df.to_dict(orient="records")
    payload = {"rows": rows}

    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        print("✅ Data successfully uploaded to Power BI Dataset.")
    else:
        print(f"❌ Failed to upload data: {response.text}")

# 🔹 Run the migration
if __name__ == "__main__":
    print("🚀 Starting Tableau to Power BI (Fabric) Migration...")
    upload_data_to_power_bi()
    print("🎉 Migration completed! Check your Power BI dataset.")
