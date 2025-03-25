import pandas as pd
import requests
import msal
from openpyxl import load_workbook
import os


class PBIManager():
    """
        A class to manage authentication and API interactions with Microsoft Power BI.

        Attributes:
            tenant_id (str): Azure Active Directory tenant ID.
            client_id (str): Application (client) ID.
            client_secret (str): Client secret for authentication.
            scope (list): API permissions required for authentication.
            access_token (str or None): Access token for API requests.
            authority (str): Authority URL for authentication.
        """
    def __init__(self,tenant_id: str, client_id: str, client_secret: str, scope: str):
        """
        Initializes the PBIManager with authentication details.

        Args:
            tenant_id (str): The Azure tenant ID.
            client_id (str): The client ID of the registered app.
            client_secret (str): The client secret for authentication.
            scope (str): The required API permission scope.
        """
        
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scope = [scope]
        self.access_token = None
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"

    def get_token(self):
        """
        Retrieves an access token using MSAL Confidential Client Application.

        Raises:
            Exception: If the token retrieval fails.
        """
        app = msal.ConfidentialClientApplication(self.client_id, authority=self.authority, client_credential=self.client_secret)

        token_response = app.acquire_token_for_client(scopes= self.scope)
        self.access_token = token_response.get('access_token')

        if not self.access_token:
            raise Exception('❌ Failed to retrieve the token')
        
    def get_list_of_datasets(self, workspace_id: str):
        """
        Fetches a list of datasets from a given Power BI workspace.

        Args:
            workspace_id (str): The ID of the Power BI workspace.

        Returns:
            pd.DataFrame: A DataFrame containing dataset IDs and names.

        Raises:
            Exception: If no access token is available.
        """
        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            data = response.json()
            datasets = data.get('value',[])

            df = pd.DataFrame(datasets, columns=['id', 'name'])
            return df 
        else:
            raise Exception(f"❌ Error: {response.status_code} - {response.text}")
        
    def get_list_of_reports(self, workspace_id: str):
        """
        Fetches a list of reports from a given Power BI workspace.

        Args:
            workspace_id (str): The ID of the Power BI workspace.

        Returns:
            pd.DataFrame: A DataFrame containing report IDs and names.

        Raises:
            Exception: If no access token is available.
        """
        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports" 
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            data = response.json()
            reports = data.get('value', [])

            df = pd.DataFrame(reports, columns=['id', 'name'])
            return df
     
        else:
            raise Exception(f"❌ Error: {response.status_code} - {response.text}")
        
    def get_reports_with_datasets(self, workspace_id: str):
        """
        Fetches a list of reports along with their associated datasets from a given Power BI workspace.

        Args:
            workspace_id (str): The ID of the Power BI workspace.

        Returns:
            pd.DataFrame: A DataFrame containing report IDs, names, dataset IDs, and dataset names.

        Raises:
            Exception: If no access token is available.
        """
        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            data = response.json()
            reports = data.get('value', [])

            report_list = []
            for report in reports:
                report_id = report.get('id')
                report_name = report.get('name')
                dataset_id = report.get('datasetId')
                
                
                dataset_name = None
                if dataset_id:
                    dataset_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}"
                    dataset_response = requests.get(dataset_url, headers=headers)
                    if dataset_response.status_code == 200:
                        dataset_data = dataset_response.json()
                        dataset_name = dataset_data.get('name')
                
                report_list.append({
                    'report_id': report_id,
                    'report_name': report_name,
                    'dataset_id': dataset_id,
                    'dataset_name': dataset_name
                })

            df = pd.DataFrame(report_list, columns=['report_id', 'report_name', 'dataset_id', 'dataset_name'])
            return df
        else:
            raise Exception(f"❌ Error: {response.status_code} - {response.text}")
        
    def execute_query(self, query_input: str, dataset_id: str):
        """
        Executes a DAX or SQL query against a Power BI dataset and returns the result as a DataFrame.

        Args:
            query_input (str): The query string to be executed.
            dataset_id (str): The ID of the Power BI dataset.

        Returns:
            pd.DataFrame: A DataFrame containing the query results.

        Raises:
            Exception: If no access token is available or if an API request fails.
        """

        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        url = f"https://api.powerbi.com/v1.0/myorg/datasets/{dataset_id}/executeQueries"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.access_token}"
        }

        query = {
            "queries": [{
                "query": f'{query_input}'
            }]
        }

        response = requests.post(url, headers=headers, json=query)

        if response.status_code == 200:
            result_json = response.json()

            try:
                data = result_json['results'][0]['tables'][0]['rows']
                df = pd.DataFrame(data)
                return df
            
            except Exception as e:
                print(f'❌ Error: {e}')
                raise
        else:
            raise Exception(f"❌ Error: {response.status_code} - {response.text}")
        
    def get_documentation(self, dataset_id, path_to_save):
        doc_var_list = ['columns','tables','measures','relations']

        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        file_name = f"Documentation_{dataset_id}.xlsx"
        file_path = os.path.join(path_to_save, file_name)
        
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            for word in doc_var_list:
                query_word = f"""
                EVALUATE
                Info_{word}
                """
                df = self.execute_query(query_input=query_word, dataset_id=dataset_id)
                df.to_excel(writer, sheet_name=word, index=False)
        print(f'✅ Documentation saved at: {file_path}')


    def refresh_dataset(self, workspace_id: str, dataset_id: str):
        """
        Triggers a refresh for a specific dataset in Power BI.
        """
        if self.access_token is None:
            raise Exception('❌ No access token, call "get_token()"')
        
        url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        response = requests.post(url, headers=headers)
        
        if response.status_code == 202:
            print(f'✅ Refresh started for dataset: {dataset_id}')
        else:
            raise Exception(f"❌ Error: {response.status_code} - {response.text}")