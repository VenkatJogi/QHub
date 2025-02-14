import requests
import pandas as pd
import xml.etree.ElementTree as ET
import csv
import io
import sys
sys.path.append(r"C:\Users\arjun\Quadrant\tableau_to_power_bi_project")
import config
import tableauserverclient as TSC

# 🔹 Tableau Server Details (Update These)
TABLEAU_SERVER = "https://10ay.online.tableau.com"
API_VERSION = "3.24"  # Update this based on your Tableau Server version
USERNAME = config.tableau_username
PASSWORD = config.tableau_password
SITE_NAME = ""  # Leave empty for default site

dashboards_query = """
    query dashboards {
      workbooks {
        name
        dashboards {
          name
          id
          sheets {
            name
          }
        }
      } 
    }
    """

worksheets_query = """
    query worksheets {
      workbooks {
        name
        sheets {
          name
          id
          upstreamFields {
            name
          }
        }
      } 
    }
    """

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
    
    # print(f"🔍 Status Code: {response.status_code}")
    # print(f"🔍 Response Text: {response.text}")

    if response.status_code != 200:
        raise Exception(f"🚨 Authentication failed! Status: {response.status_code}\n{response.text}")

    root = ET.fromstring(response.text)
    namespaces = {"t": "http://tableau.com/api"}  # Namespace correction

    token_elem = root.find(".//t:credentials", namespaces)
    if token_elem is None:
        raise Exception("🚨 Authentication failed! No credentials found in response.")

    token = token_elem.get("token")
    site_id = root.find(".//t:site", namespaces).get("id")

    # print(f"✅ Authenticated! Token: {token}")
    return token, site_id

#🔹 Fetch dashboards
def get_dashboards(response, workbook_names):
    dashboards_result = []
    workbooks = response['workbooks']
    for workbook in workbooks:
        workbook_name = workbook['name']
        if (workbook_name not in workbook_names):
            continue
        dashboards = workbook['dashboards']
        for dashboard in dashboards:
            dashboard_name = dashboard['name']
            dashboard_id = dashboard['id']
            worksheets = []
            for item in dashboard['sheets']:
                worksheets.append(item['name'])
            worksheets = ", ".join(worksheets)  
            dashboards_result.append({
                "Workbook Name": workbook_name,
                "Dashboard Name": dashboard_name,
                "Dashboard ID": dashboard_id,
                "Sheets Used": worksheets
            })     
    return dashboards_result

# 🔹 Fetch worksheets
def get_worksheets(response, workbook_names):
    worksheets_result = []
    workbooks = response['workbooks']
    for workbook in workbooks:
        workbook_name = workbook['name']
        if (workbook_name not in workbook_names):
            continue
        worksheets = workbook['sheets']
        for worksheet in worksheets:
            worksheet_name = worksheet['name']
            worksheet_id = worksheet['id']
            columns = []
            for item in worksheet['upstreamFields']:
                columns.append(item['name'])
            columns = ", ".join(columns)  
            worksheets_result.append({
                "Workbook Name": workbook_name,
                "Worksheet Name": worksheet_name,
                "Worksheet ID": worksheet_id,
                "Columns Used": columns
            }) 
    return worksheets_result

# 🔹 Extract and Save Metadata
def extract_metadata():
    auth_token, site_id = get_auth_token()
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/workbooks"
    headers = {"X-Tableau-Auth": auth_token}
    response = requests.get(url, headers=headers)
    root = ET.fromstring(response.text)
    workbooks = root.findall(".//t:workbook", namespaces={"t": "http://tableau.com/api"})
    workbook_names = []
    for workbook in workbooks:
        workbook_names.append(workbook.get("name"))
    # dashboards = get_dashboards(auth_token, site_id)
    # df_dashboards = pd.DataFrame(dashboards)
    # with pd.ExcelWriter("tableau_metadata.xlsx") as writer:
    #     df_dashboards.to_excel(writer, sheet_name="Dashboards", index=False)
    # print("✅ Extracted dashboard metadata saved to 'tableau_metadata.xlsx'.")

    # Try using the Tableau Metadata API
    tableau_auth = TSC.TableauAuth(USERNAME, PASSWORD, SITE_NAME)
    server = TSC.Server(TABLEAU_SERVER)
    server.version = API_VERSION

    with server.auth.sign_in(tableau_auth):
        #Query the Metadata API
        response = server.metadata.query(dashboards_query)['data']
        dashboards = get_dashboards(response, workbook_names)
        df_dashboards = pd.DataFrame(dashboards)

        response = server.metadata.query(worksheets_query)['data']
        worksheets = get_worksheets(response, workbook_names)
        df_worksheets = pd.DataFrame(worksheets)

        with pd.ExcelWriter("tableau_metadata.xlsx") as writer:
            df_dashboards.to_excel(writer, sheet_name="Dashboards", index=False)
            df_worksheets.to_excel(writer, sheet_name="Worksheets", index=False)

# 🔹 Run the Extraction
if __name__ == "__main__":
    extract_metadata()
