from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import tempfile
import os

class SharePointConnector:
    def __init__(self, sharepoint_url: str, site_name: str, username: str, password: str):
        """Initialize SharePoint connector with credentials"""
        self.sharepoint_url = sharepoint_url
        self.site_name = site_name
        self.ctx = self._get_context(username, password)
        
    def _get_context(self, username: str, password: str) -> ClientContext:
        """Create SharePoint context with user credentials"""
        site_url = f"{self.sharepoint_url}/sites/{self.site_name}"
        credentials = UserCredential(username, password)
        return ClientContext(site_url).with_credentials(credentials)
    
    def get_all_documents(self):
        """Retrieve all documents from the SharePoint site"""
        try:
            # Get the Shared Documents library
            doc_lib = self.ctx.web.lists.get_by_title("Documents")
            items = doc_lib.items.select(["FileLeafRef", "File"]).get().execute_query()
            
            documents = []
            for item in items:
                if hasattr(item, "File"):
                    file_name = item.file_leaf_ref
                    file_url = item.file.serverRelativeUrl
                    
                    # Download file to temp directory
                    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                        file = self.ctx.web.get_file_by_server_relative_url(file_url)
                        file.download(temp_file.name).execute_query()
                        
                        # Read file content
                        with open(temp_file.name, "r", encoding="utf-8") as f:
                            content = f.read()
                        
                        documents.append({
                            "name": file_name,
                            "content": content,
                            "url": file_url
                        })
                        
                        # Clean up temp file
                        os.unlink(temp_file.name)
            
            return documents
        
        except Exception as e:
            raise Exception(f"Error retrieving documents: {str(e)}")
    
    def get_document_by_name(self, document_name: str):
        """Retrieve a specific document by name"""
        try:
            doc_lib = self.ctx.web.lists.get_by_title("Documents")
            items = doc_lib.items.filter(f"FileLeafRef eq '{document_name}'").get().execute_query()
            
            if len(items) > 0:
                item = items[0]
                file_url = item.file.serverRelativeUrl
                
                with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                    file = self.ctx.web.get_file_by_server_relative_url(file_url)
                    file.download(temp_file.name).execute_query()
                    
                    with open(temp_file.name, "r", encoding="utf-8") as f:
                        content = f.read()
                    
                    # Clean up temp file
                    os.unlink(temp_file.name)
                    
                    return {
                        "name": document_name,
                        "content": content,
                        "url": file_url
                    }
            
            return None
            
        except Exception as e:
            raise Exception(f"Error retrieving document: {str(e)}")
