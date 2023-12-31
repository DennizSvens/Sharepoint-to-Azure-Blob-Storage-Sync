# Sharepoint to Azure Blob Storage Sync Script

This script is intended to synchronize files from Sharepoint to Azure Blob Storage. It recursively checks for file changes in a specified Sharepoint folder and reflects those changes in Azure Blob Storage.


## Important Notes

- This script was designed for efficient synchronization and includes multi-threading support to improve performance. However, it's advisable to test thoroughly in a non-production environment before deploying.
- The script requires an empty folder/map in the target Azure container. It writes `sp_last_modified` to the blob's metadata to track changes since Sharepoint does not appear to expose MD5 for data integrity checks.
- To sync multiple folders, utilize the config file (refer to config.example.json for an example setup) and set CONFIG_FILE in .env.
- This script has been tested for use with serverless [Azure Function Apps](https://learn.microsoft.com/en-us/azure/azure-functions/functions-reference-python?tabs=asgi%2Capplication-level&pivots=python-mode-decorators) running Python, utilizing timer triggers for automatic synchronization.

## Prerequisites

- Python 3
- A `.env` file based on the `.env.example` provided
- Packages specified in `requirements.txt`
- Azure AD app with certificate authentication (see [this guide](https://github.com/vgrem/Office365-REST-Python-Client/wiki/How-to-connect-to-SharePoint-Online-with-certificate-credentials
) for more information)
- Azure Storage Account with a container and folder for uploads

## Installation

1. Clone this repository:

    ```bash
    git clone <repository_url>
    ```

2. Navigate to the cloned directory:

    ```bash
    cd path/to/cloned/directory
    ```

3. Install the required packages:

    ```bash
    pip install -r requirements.txt
    ```

4. Copy the `.env.example` file to `.env`:

    ```bash
    cp .env.example .env
    ```

5. Update the `.env` file with your specific configurations.

## Usage

1. Run the sync script:

    ```bash
    python sync.py
    ```

    This script will detect changes between Sharepoint and Azure Blob Storage. If `DRY_RUN` is set to `True` in `.env`, it will only print the changes without executing them. Otherwise, it will perform the necessary uploads, updates, or deletions.

## Configuration

The following environment variables are essential for the script's functioning:

- **AZURE_AD_CLIENT_ID**: Client ID from Azure AD App.
- **AZURE_AD_TENANT_ID**: Tenant ID from Azure AD.
- **AZURE_AD_CERTIFICATE_NAME**: File name of the certificate. This should be uploaded to Azure AD and must be in the same directory as this script.
- **AZURE_AD_CERTIFICATE_THUMBPRINT**: Thumbprint of the certificate uploaded to Azure AD.
- **AZURE_STORAGE_CONNECTION_STRING**: Connection string to the Azure Storage Account.
- **AZURE_STORAGE_CONTAINER_NAME**: Container name in Azure Storage Account where files should be uploaded.
- **AZURE_STORAGE_FOLDER_NAME**: Folder name in Azure Storage Account for file uploads.
- **SHAREPOINT_BASE**: The base URL for your Sharepoint (e.g., `https://tenant.sharepoint.com`).
- **SHAREPOINT_SITE**: The specific Sharepoint site you want to synchronize (e.g., `/sites/yoursharepointsite/`).
- **SHAREPOINT_TARGET_FOLDER**: The target folder in Sharepoint to sync (e.g., `Shared Documents/Folder1`).
- **DRY_RUN**: If set to `True`, only the changes that would occur are printed without actual execution.
- **CONFIG_FILE**: If you need to sync multiple folders you can set this path to the config file that contains an array of AZURE_STORAGE_CONTAINER_NAME, AZURE_STORAGE_FOLDER_NAME, SHAREPOINT_SITE ,SHAREPOINT_TARGET_FOLDER
- **MAX_WORKERS**: The maximum number of threads to use for uploads. Defaults to 1. Increase this number to improve performance. Optimal performance is typically seen with 5-10 threads.

## Contributions

Contributions to improve this script are welcome. Please fork the repository, make your changes, and submit a pull request.