# Power BI API Integration

This repository provides a Python class (`PBIManager`) for interacting with the Microsoft Power BI API. It handles authentication using MSAL and offers methods to fetch datasets, reports, execute queries, and generate dataset documentation.

## Key Features

- **Authentication**: Uses MSAL to obtain access tokens for API requests.
- **Dataset Management**: Retrieves lists of datasets from specified workspaces.
- **Report Management**: Fetches lists of reports, including those with associated datasets.
- **Query Execution**: Executes DAX or SQL queries against Power BI datasets.
- **Documentation**: Generates Excel documentation for dataset structures.

## Usage

1. **Initialization**: Initialize the `PBIManager` class with your Azure tenant ID, client ID, client secret, and scope.
2. **Get Token**: Call `get_token()` to obtain an access token.
3. **Fetch Datasets**: Use `get_list_of_datasets(workspace_id)` to retrieve datasets.
4. **Fetch Reports**: Use `get_list_of_reports(workspace_id)` or `get_reports_with_datasets(workspace_id)` to retrieve reports.
5. **Execute Query**: Use `execute_query(query_input, dataset_id)` to execute a query.
6. **Generate Documentation**: Use `get_documentation(dataset_id, path_to_save)` to generate dataset documentation.

### Example
```python
from classes_pbi_api import PBIManager
```

**Initialize PBIManager**
```python
pbi_manager = PBIManager(
tenant_id='Your tenant ID',
client_id='Your client ID',
client_secret='Your client secret',
scope='https://analysis.windows.net/powerbi/api/.default'
)
```
**Get access token**
```python
pbi_manager.get_token()
```
**Fetch datasets**
```python
df_datasets = pbi_manager.get_list_of_datasets(workspace_id='Your workspace ID')
```
**Fetch reports**
```python
df_reports = pbi_manager.get_list_of_reports(workspace_id='Your workspace ID')
```
**Execute query**
```python
query_result = """EVALUATE YOUR DAX QUERY"""
df_query = pbi_manager.execute_query(query_input=query_result, dataset_id='Your Dataset ID')
```
**Generate documentation**
```python
pbi_manager.get_documentation(dataset_id='Your Dataset ID', path_to_save=r'Your save path')
```

