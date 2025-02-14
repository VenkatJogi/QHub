import requests
import pandas as pd
import xml.etree.ElementTree as ET
import csv
import io
import config

# 🔹 Tableau Server Details (Update These)
TABLEAU_SERVER = "https://10ay.online.tableau.com"
API_VERSION = "3.24"  # Update this based on your Tableau Server version
USERNAME = config.tableau_username
PASSWORD = config.tableau_password
SITE_NAME = ""  # Leave empty for default site

# 🔹 Authentication Function
def get_auth_token():
    """Authenticate with Tableau Server and return an auth token."""
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/auth/signin"
    
    payload = f"""
        <tsRequest>
            <credentials name="{USERNAME}" password="{PASSWORD}">
                <site contentUrl="{SITE_NAME}"/>
            </credentials>
        </tsRequest>
    """
    headers = {"Content-Type": "application/xml"}

    response = requests.post(url, data=payload, headers=headers)
    
    print(f"🔍 Status Code: {response.status_code}")
    print(f"🔍 Response Text: {response.text}")

    if response.status_code != 200:
        raise Exception(f"🚨 Authentication failed! Status: {response.status_code}\n{response.text}")

    root = ET.fromstring(response.text)
    namespaces = {"t": "http://tableau.com/api"}  # Namespace correction

    token_elem = root.find(".//t:credentials", namespaces)
    if token_elem is None:
        raise Exception("🚨 Authentication failed! No credentials found in response.")

    token = token_elem.get("token")
    site_id = root.find(".//t:site", namespaces).get("id")

    print(f"✅ Authenticated! Token: {token}")
    return token, site_id

# 🔹 Fetch Workbooks (to find dashboards)
def get_dashboards(auth_token, site_id):
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/workbooks"
    headers = {"X-Tableau-Auth": auth_token}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        root = ET.fromstring(response.text)
        dashboards = []
        for workbook in root.findall(".//t:workbook", namespaces={"t": "http://tableau.com/api"}):
            workbook_id = workbook.get("id")
            workbook_name = workbook.get("name")
            dashboards.extend(get_sheets(auth_token, site_id, workbook_id, workbook_name))
        return dashboards
    else:
        raise Exception(f"🚨 Failed to fetch workbooks! Error: {response.text}")

# 🔹 Fetch Sheets (Dashboards)
def get_sheets(auth_token, site_id, workbook_id, workbook_name):
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/workbooks/{workbook_id}/views"
    headers = {"X-Tableau-Auth": auth_token}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        root = ET.fromstring(response.text)
        sheets = []
        for sheet in root.findall(".//t:view", namespaces={"t": "http://tableau.com/api"}):
            sheet_name = sheet.get("name")
            sheet_id = sheet.get("id")
            content_url = sheet.get("contentUrl")
            visual_type, columns = get_visual_metadata(auth_token, site_id, sheet_id)
            sheets.append({
                "Workbook Name": workbook_name,
                "Dashboard Name": sheet_name,
                "Dashboard ID": sheet_id,
                "Dashboard URL": f"{TABLEAU_SERVER}/#/site/{site_id}/views/{content_url}",
                "Visual Type": visual_type,
                "Columns Used": ", ".join(columns)
            })
        return sheets
    else:
        return []

# 🔹 Fetch Visual Type & Columns
def get_visual_metadata(auth_token, site_id, view_id):
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/views/{view_id}/data"
    headers = {"X-Tableau-Auth": auth_token}

    response = requests.get(url, headers=headers)

    print(f"🔍 Response Status: {response.status_code}")
    print(f"🔍 Response Text Preview: {response.text[:200]}")  # Show only the first 200 characters

    if response.status_code == 200:
        # 🔹 Check if response is CSV (contains commas in the first line)
        if "," in response.text[:100]:  
            csv_data = io.StringIO(response.text)
            reader = csv.reader(csv_data)
            columns = next(reader)  # First row is column headers
            
            return "Table", columns  # Assuming all CSV responses represent table-type visuals

        else:
            print("❌ Unexpected response format (not CSV or XML).")
            return "Unknown", []

    return "Unknown", []


# 🔹 Extract and Save Metadata
def extract_metadata():
    auth_token, site_id = get_auth_token()
    dashboards = get_dashboards(auth_token, site_id)
    df_dashboards = pd.DataFrame(dashboards)
    with pd.ExcelWriter("tableau_metadata5.xlsx") as writer:
        df_dashboards.to_excel(writer, sheet_name="Dashboards", index=False)
    print("✅ Extracted dashboard metadata saved to 'tableau_metadata.xlsx'.")

# 🔹 Run the Extraction
if __name__ == "__main__":
    extract_metadata()
