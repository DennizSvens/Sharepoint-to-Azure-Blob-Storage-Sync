import os
from office365.sharepoint.client_context import ClientContext
from azure.storage.blob import BlobServiceClient
from loguru import logger
from decouple import config
import pandas as pd


# Load environment variables
AZURE_AD_CLIENT_ID = config("AZURE_AD_CLIENT_ID")
AZURE_AD_TENANT_ID = config("AZURE_AD_TENANT_ID")
AZURE_AD_CERTIFICATE_PATH = "{0}/{1}".format(os.path.dirname(__file__), config("AZURE_AD_CERTIFICATE_NAME"))
AZURE_AD_CERTIFICATE_THUMBPRINT = config("AZURE_AD_CERTIFICATE_THUMBPRINT")
AZURE_STORAGE_CONNECTION_STRING = config("AZURE_STORAGE_CONNECTION_STRING")
AZURE_STORAGE_CONTAINER_NAME = config("AZURE_STORAGE_CONTAINER_NAME")
AZURE_STORAGE_FOLDER_NAME = config("AZURE_STORAGE_FOLDER_NAME")
SHAREPOINT_BASE = config("SHAREPOINT_BASE")
SHAREPOINT_SITE = config("SHAREPOINT_SITE")
SHAREPOINT_TARGET_FOLDER = config("SHAREPOINT_TARGET_FOLDER")
DRY_RUN = config("DRY_RUN", cast=bool)


# Constants
RECURSIVE = True
UPLOAD = "UPLOAD"
UPDATE = "UPDATE"
DELETE = "DELETE"


class File:
    def __init__(self, path, target):
        self.path = path
        self.target = target
    def __str__(self):
        return self.path
    def __repr__(self):
        return self.__str__()

    
class SharepointFile (File):
    def __init__(self, path, target):
        super().__init__(path, target)
        logger.debug("Sharepoint File initialized: {0}".format(path))
    def __str__(self):
        return self.folder_path()
    def __repr__(self):
        return self.__str__()
    def folder_path(self):
        return self.path.replace(SHAREPOINT_TARGET_FOLDER, "").replace(SHAREPOINT_SITE,"")
    def get_binary_stream(self):
        file = self.target.get().execute_query()
        return file.read()
    def azure_target_path(self):
        return AZURE_STORAGE_FOLDER_NAME + self.folder_path()
    def get_modified_date(self):
        return self.target.properties['TimeLastModified'].strftime('%Y-%m-%d %H:%M:%S')
    def upload_to_blob(self, container_client, overwrite=False):  # Pass the container_client here
        container_client.get_blob_client(self.azure_target_path()).upload_blob(data=self.get_binary_stream(), metadata={"sp_last_modified": self.get_modified_date()}, overwrite=overwrite)

class AzureFile (File): 
    def __init__(self, path, target):
        super().__init__(path, target)
        logger.debug("Azure File initialized: {0}".format(path))
    def __str__(self):
        return self.path
    def __repr__(self):
        return self.__str__()
    def get_modified_date(self):
        return self.target.metadata["sp_last_modified"]
    def delete_blob(self, container_client):  # Pass the container_client here
        container_client.get_blob_client(self.path).delete_blob()



class SyncManager:
    def __init__(self):
        self.ctx = None
        self.blob_service_client = None
        self.container_client = None
        self.connect_to_sharepoint()
        self.connect_to_azure()

    def connect_to_sharepoint(self):
        try:
            self.ctx = ClientContext(SHAREPOINT_BASE + SHAREPOINT_SITE).with_client_certificate(**{
                "tenant": AZURE_AD_TENANT_ID,
                "client_id": AZURE_AD_CLIENT_ID,
                "cert_path": AZURE_AD_CERTIFICATE_PATH,
                "thumbprint": AZURE_AD_CERTIFICATE_THUMBPRINT
            })
            target_web = self.ctx.web.get().execute_query()
            logger.info("Connected to Sharepoint: {0}".format(target_web.url))
        except Exception as error: 
            logger.error("Could not connect to Sharepoint: {0}".format(error))
            exit(1)

    def connect_to_azure(self):
        try:
            self.blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)
            self.container_client = self.blob_service_client.get_container_client(AZURE_STORAGE_CONTAINER_NAME)
            container_exists = self.container_client.exists()
            if container_exists:
                logger.info("Connected to Azure: {0}".format(self.container_client.url))
            else:
                logger.error("Container does not exist!")
                exit(1)
        except Exception as error:
            logger.error("Could not connect to Azure: {0}".format(error))
            exit(1)
    def get_sharepoint_files_recursive(self, relative_folder_url):
      logger.debug("Retrieving files in: {0}".format(relative_folder_url))
      relative_file_url_arr = []
      folder = self.ctx.web.get_folder_by_server_relative_url(relative_folder_url) 
      files = folder.files
      self.ctx.load(files)
      self.ctx.execute_query()
      for file in files:
          sharepoint_file = SharepointFile(file.properties["ServerRelativeUrl"], file)
          relative_file_url_arr.append(sharepoint_file)
      if RECURSIVE:
          folders = folder.folders
          self.ctx.load(folders)
          self.ctx.execute_query()
          for folder in folders:
              relative_file_url_arr += self.get_sharepoint_files_recursive(folder.properties["ServerRelativeUrl"])
      return relative_file_url_arr


    def get_azure_files_recursive(self, relative_folder_url):
      logger.debug("Retrieving files in: {0}".format(relative_folder_url))
      relative_file_url_arr = []
      blob_list = self.container_client.list_blobs(name_starts_with=relative_folder_url, include="metadata")
      for blob in blob_list:
          azure_file = AzureFile(blob.name, blob)
          relative_file_url_arr.append(azure_file)

      return relative_file_url_arr
    
    def detect_changes(self):
        sharepoint_folder_url = SHAREPOINT_SITE + SHAREPOINT_TARGET_FOLDER
        sharepoint_files = self.get_sharepoint_files_recursive(sharepoint_folder_url)
        azure_files = self.get_azure_files_recursive(AZURE_STORAGE_FOLDER_NAME)

        changes = []
        for sharepoint_file in sharepoint_files:
            azure_file = next((af for af in azure_files if af.path == sharepoint_file.azure_target_path()), None)
            if not azure_file:
                changes.append({"OPERATION": UPLOAD, "SOURCE": sharepoint_file, "TARGET": sharepoint_file.azure_target_path(), "FILE": sharepoint_file})
            elif azure_file.get_modified_date() != sharepoint_file.get_modified_date():
                changes.append({"OPERATION": UPDATE, "SOURCE": sharepoint_file, "TARGET": sharepoint_file.azure_target_path(), "FILE": sharepoint_file})

        for azure_file in azure_files:
            if not any(sp_file.azure_target_path() == azure_file.path for sp_file in sharepoint_files):
                changes.append({"OPERATION": DELETE, "SOURCE": azure_file, "TARGET": azure_file, "FILE": azure_file})

        return changes
    def execute_changes(self, changes):
        for change in changes:
            if change["OPERATION"] == UPLOAD:
                logger.info("Uploading: {0}".format(change["SOURCE"]))
                change["FILE"].upload_to_blob(self.container_client)
            elif change["OPERATION"] == UPDATE:
                logger.info("Updating: {0}".format(change["SOURCE"]))
                change["FILE"].upload_to_blob(self.container_client, overwrite=True)
            elif change["OPERATION"] == DELETE:
                logger.info("Deleting: {0}".format(change["SOURCE"]))
                change["FILE"].delete_blob(self.container_client)

    def print_changes(self, changes):
        df = pd.DataFrame(changes)
        print(df.to_string(index=False))

manager = SyncManager()
changes = manager.detect_changes()
manager.print_changes(changes)
if not DRY_RUN and len(changes) > 0:
    logger.info("Executing changes...")
    manager.execute_changes(changes)
    logger.info("Changes executed.")
else:
    logger.info("No changes to execute.")