from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
from google.oauth2 import service_account
from googleapiclient.discovery import build

def init_service():
    """Initialize Google Drive service"""
    SCOPES = ['https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = r'Auto Slides\key.json'
    
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    service = build('drive', 'v3', credentials=credentials)
    return service

_service = None

def get_service():
    """Get or create Google Drive service"""
    global _service
    if _service is None:
        _service = init_service()
    return _service

def create_folder(folder_name: str, parent_folder_id: str = None) -> str:
    """
    Create a folder in Google Drive
    Args:
        folder_name: Name of the folder to create
        parent_folder_id: Optional parent folder ID
    Returns:
        str: ID of created folder
    """
    service = get_service()
    
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    
    if parent_folder_id:
        file_metadata['parents'] = [parent_folder_id]
    
    folder = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    
    return folder.get('id')

def search_files(parent_folder_id: str = None, file_type: str = None, additional_filter: str = None) -> list:
    """
    Search for files/folders in Google Drive
    Args:
        parent_folder_id: Optional parent folder ID to search in
        file_type: Optional type ('file' or 'folder')
        additional_filter: Optional additional query filter
    Returns:
        list: List of matching files/folders
    """
    service = get_service()
    
    query_parts = []
    
    if parent_folder_id:
        query_parts.append(f"'{parent_folder_id}' in parents")
    
    if file_type == 'folder':
        query_parts.append("mimeType='application/vnd.google-apps.folder'")
    elif file_type == 'file':
        query_parts.append("mimeType!='application/vnd.google-apps.folder'")
    
    if additional_filter:
        query_parts.append(additional_filter)
    
    query_parts.append("trashed=false")
    
    query = " and ".join(query_parts)
    
    results = service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name, mimeType)',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True
    ).execute()
    
    return results.get('files', [])

def delete_file(file_id: str) -> bool:
    """
    Delete a file/folder from Google Drive
    Args:
        file_id: ID of the file/folder to delete
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        service = get_service()
        service.files().delete(fileId=file_id).execute()
        return True
    except Exception as e:
        print(f"Error deleting file {file_id}: {str(e)}")
        return False

def upload_file(file_obj, filename: str, parent_folder_id: str = None, file_type: str = None) -> str:
    """
    Upload a file to Google Drive
    Args:
        file_obj: File object or path to upload
        filename: Name to give the file in Drive
        parent_folder_id: Optional parent folder ID
        file_type: Optional file type (e.g., 'pdf', 'pptx')
    Returns:
        str: ID of uploaded file
    """
    service = get_service()
    
    mime_types = {
        'pdf': 'application/pdf',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    }
    
    mime_type = mime_types.get(file_type, 'application/octet-stream')
    
    file_metadata = {'name': filename}
    if parent_folder_id:
        file_metadata['parents'] = [parent_folder_id]
    
    # Handle both file paths and file objects
    if isinstance(file_obj, str):
        media = MediaFileUpload(
            file_obj,
            mimetype=mime_type,
            resumable=True
        )
    else:
        # For file objects, wrap them in a BytesIO buffer
        import io
        media = MediaIoBaseUpload(
            io.BytesIO(file_obj.read()),
            mimetype=mime_type,
            resumable=True
        )
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()
    
    return file.get('id') 