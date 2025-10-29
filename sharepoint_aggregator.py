"""
Folder Aggregator

This tool navigates through folders in the same location as the script,
processes Excel questionnaires, and aggregates the data into a single Excel file.
It supports both local filesystem and SharePoint modes.
"""

import os
import sys
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Union
from dotenv import load_dotenv
import openpyxl
import pandas as pd

# Optional SharePoint imports
try:
    from office365.runtime.auth.authentication_context import AuthenticationContext
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    SHAREPOINT_AVAILABLE = True
except ImportError:
    SHAREPOINT_AVAILABLE = False


class LocalFolderProcessor:
    """Handles local filesystem operations for folder and file scanning."""
    
    def __init__(self, base_path: Union[str, Path]):
        """
        Initialize local folder processor.
        
        Args:
            base_path: Base directory path to scan for folders
        """
        self.base_path = Path(base_path).resolve()
        
    def get_folders(self) -> List[Path]:
        """
        Get all folders in the base path.
        
        Returns:
            List of folder paths
        """
        try:
            folders = [f for f in self.base_path.iterdir() if f.is_dir() and not f.name.startswith('.')]
            return folders
        except Exception as e:
            print(f"Error getting folders from {self.base_path}: {e}")
            return []
    
    def get_excel_files_in_folder(self, folder_path: Path) -> List[Path]:
        """
        Get all Excel files in a folder.
        
        Args:
            folder_path: Path to the folder
            
        Returns:
            List of Excel file paths
        """
        try:
            excel_files = []
            for file in folder_path.iterdir():
                if file.is_file() and file.suffix.lower() in ['.xlsx', '.xls'] and not file.name.startswith('~$'):
                    excel_files.append(file)
            return excel_files
        except Exception as e:
            print(f"Error getting files from {folder_path}: {e}")
            return []


class SharePointConnector:
    """Handles SharePoint authentication and file operations."""
    
    def __init__(self, site_url: str, username: str, password: str):
        """
        Initialize SharePoint connector.
        
        Args:
            site_url: SharePoint site URL
            username: SharePoint username
            password: SharePoint password
        """
        self.site_url = site_url
        self.username = username
        self.password = password
        self.ctx = None
        
    def connect(self) -> bool:
        """
        Authenticate and connect to SharePoint.
        
        Returns:
            True if connection successful, False otherwise
        """
        try:
            auth_ctx = AuthenticationContext(self.site_url)
            if auth_ctx.acquire_token_for_user(self.username, self.password):
                self.ctx = ClientContext(self.site_url, auth_ctx)
                # Test the connection
                web = self.ctx.web
                self.ctx.load(web)
                self.ctx.execute_query()
                print(f"Connected to SharePoint site: {web.properties['Title']}")
                return True
            else:
                print("Authentication failed")
                return False
        except Exception as e:
            print(f"Error connecting to SharePoint: {e}")
            return False
    
    def get_folders_in_path(self, folder_path: str) -> List[Any]:
        """
        Get all folders within a specified path.
        
        Args:
            folder_path: Path to the folder in SharePoint
            
        Returns:
            List of folder objects
        """
        try:
            # Get the folder
            folder = self.ctx.web.get_folder_by_server_relative_url(folder_path)
            folders = folder.folders
            self.ctx.load(folders)
            self.ctx.execute_query()
            return [f for f in folders]
        except Exception as e:
            print(f"Error getting folders from {folder_path}: {e}")
            return []
    
    def get_excel_files_in_folder(self, folder_url: str) -> List[Any]:
        """
        Get all Excel files in a folder.
        
        Args:
            folder_url: Server relative URL of the folder
            
        Returns:
            List of Excel file objects
        """
        try:
            folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)
            files = folder.files
            self.ctx.load(files)
            self.ctx.execute_query()
            
            excel_files = []
            for file in files:
                file_name = file.properties.get('Name', '')
                if file_name.endswith(('.xlsx', '.xls')) and not file_name.startswith('~$'):
                    excel_files.append(file)
            
            return excel_files
        except Exception as e:
            print(f"Error getting files from {folder_url}: {e}")
            return []
    
    def download_file(self, file_url: str, local_path: str) -> bool:
        """
        Download a file from SharePoint.
        
        Args:
            file_url: Server relative URL of the file
            local_path: Local path where to save the file
            
        Returns:
            True if download successful, False otherwise
        """
        try:
            response = File.open_binary(self.ctx, file_url)
            with open(local_path, 'wb') as local_file:
                local_file.write(response.content)
            return True
        except Exception as e:
            print(f"Error downloading file {file_url}: {e}")
            return False


class QuestionnaireExtractor:
    """Extracts data from Excel questionnaire files."""
    
    @staticmethod
    def extract_data(file_path: str, source_folder: str) -> Dict[str, Any]:
        """
        Extract data from a questionnaire Excel file.
        
        Args:
            file_path: Path to the Excel file
            source_folder: Name of the source folder
            
        Returns:
            Dictionary containing extracted data
        """
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            sheet = workbook.active
            
            # Extract metadata from row 1
            application = sheet['B1'].value
            responsible = sheet['C1'].value
            deputy = sheet['D1'].value
            
            # Extract questions, answers (column C), and comments (column D) from rows 3-19
            questions_answers = []
            for row in range(3, 20):  # Rows 3 to 19 inclusive
                question = sheet[f'B{row}'].value
                answer = sheet[f'C{row}'].value
                comment = sheet[f'D{row}'].value
                
                questions_answers.append({
                    'question': question,
                    'answer': answer,
                    'comment': comment
                })
            
            workbook.close()
            
            return {
                'source_folder': source_folder,
                'application': application,
                'responsible': responsible,
                'deputy': deputy,
                'questions_answers': questions_answers
            }
        except Exception as e:
            print(f"Error extracting data from {file_path}: {e}")
            return None


class DataAggregator:
    """Aggregates data from multiple questionnaires into a single Excel file."""
    
    def __init__(self):
        """Initialize the aggregator."""
        self.all_data = []
    
    def add_questionnaire(self, data: Dict[str, Any]):
        """
        Add questionnaire data to the aggregation.
        
        Args:
            data: Dictionary containing questionnaire data
        """
        if data:
            self.all_data.append(data)
    
    def export_to_excel(self, output_file: str):
        """
        Export aggregated data to Excel.
        
        Args:
            output_file: Path to the output Excel file
        """
        if not self.all_data:
            print("No data to export")
            return
        
        # Prepare data for DataFrame
        rows = []
        for data in self.all_data:
            # Create a base row with metadata
            row = {
                'Application': data['application'],
                'Answered': data['responsible'],
                'App Responsible': data['responsible'],
                'Deputy': data['deputy']
            }
            
            # Add answers (Q1-Q17)
            for i, qa in enumerate(data['questions_answers'], start=1):
                row[f'Q{i}'] = qa['answer']
            
            # Add comments (COMM Q1-COMM Q17)
            for i, qa in enumerate(data['questions_answers'], start=1):
                row[f'COMM Q{i}'] = qa['comment']
            
            rows.append(row)
        
        # Create DataFrame and export
        df = pd.DataFrame(rows)
        df.to_excel(output_file, index=False)
        print(f"Aggregated data exported to: {output_file}")


def main():
    """Main function to orchestrate the questionnaire aggregation."""
    # Load environment variables from .env file
    load_dotenv()
    
    # Get configuration from environment
    mode = os.getenv('MODE', 'local').lower()
    output_file = os.getenv('OUTPUT_FILE', 'aggregated_questionnaires.xlsx')
    
    print("Questionnaire Aggregator")
    print("=" * 50)
    
    if mode == 'local':
        # Local filesystem mode
        print("Mode: Local Filesystem")
        
        # Get the base path: use BASE_PATH from env, or current working directory
        base_path = os.getenv('BASE_PATH')
        if base_path:
            base_path = Path(base_path).resolve()
        else:
            # Use current working directory (where the user runs the script from)
            base_path = Path.cwd()
        
        print(f"Scanning folders in: {base_path}")
        
        # Initialize local processor
        processor = LocalFolderProcessor(base_path)
        
        # Initialize aggregator
        aggregator = DataAggregator()
        
        # Get all folders in the base path
        folders = processor.get_folders()
        print(f"Found {len(folders)} folders")
        
        # Process each folder
        for folder in folders:
            folder_name = folder.name
            
            print(f"\nProcessing folder: {folder_name}")
            
            # Get Excel files in this folder
            excel_files = processor.get_excel_files_in_folder(folder)
            print(f"  Found {len(excel_files)} Excel file(s)")
            
            # Process each Excel file
            for excel_file in excel_files:
                file_name = excel_file.name
                
                print(f"    Processing: {file_name}")
                
                # Extract data
                data = QuestionnaireExtractor.extract_data(str(excel_file), folder_name)
                if data:
                    aggregator.add_questionnaire(data)
                    print(f"      Extracted data from {file_name}")
                else:
                    print(f"      Failed to extract data from {file_name}")
        
        # Export aggregated data
        print("\n" + "=" * 50)
        aggregator.export_to_excel(output_file)
        print(f"\nProcessing complete! Total questionnaires processed: {len(aggregator.all_data)}")
        
    elif mode == 'sharepoint':
        # SharePoint mode
        if not SHAREPOINT_AVAILABLE:
            print("Error: SharePoint mode requires Office365-REST-Python-Client package")
            print("Install it with: pip install Office365-REST-Python-Client")
            sys.exit(1)
        
        print("Mode: SharePoint")
        
        site_url = os.getenv('SHAREPOINT_SITE_URL')
        username = os.getenv('SHAREPOINT_USERNAME')
        password = os.getenv('SHAREPOINT_PASSWORD')
        folder_path = os.getenv('QUESTIONNAIRE_FOLDER_PATH')
        
        # Validate configuration
        if not all([site_url, username, password, folder_path]):
            print("Error: Missing required configuration for SharePoint mode. Please check your .env file.")
            print("Required variables: SHAREPOINT_SITE_URL, SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD, QUESTIONNAIRE_FOLDER_PATH")
            sys.exit(1)
        
        # Connect to SharePoint
        connector = SharePointConnector(site_url, username, password)
        if not connector.connect():
            print("Failed to connect to SharePoint")
            sys.exit(1)
        
        # Initialize aggregator
        aggregator = DataAggregator()
        
        # Create temporary directory for downloads
        with tempfile.TemporaryDirectory() as temp_dir:
            print(f"\nScanning folder: {folder_path}")
            
            # Get all folders in the questionnaire path
            folders = connector.get_folders_in_path(folder_path)
            print(f"Found {len(folders)} folders")
            
            # Process each folder
            for folder in folders:
                folder_name = folder.properties.get('Name', '')
                folder_url = folder.properties.get('ServerRelativeUrl', '')
                
                print(f"\nProcessing folder: {folder_name}")
                
                # Get Excel files in this folder
                excel_files = connector.get_excel_files_in_folder(folder_url)
                print(f"  Found {len(excel_files)} Excel file(s)")
                
                # Download and process each Excel file
                for excel_file in excel_files:
                    file_name = excel_file.properties.get('Name', '')
                    file_url = excel_file.properties.get('ServerRelativeUrl', '')
                    
                    print(f"    Processing: {file_name}")
                    
                    # Download file
                    local_path = os.path.join(temp_dir, file_name)
                    if connector.download_file(file_url, local_path):
                        # Extract data
                        data = QuestionnaireExtractor.extract_data(local_path, folder_name)
                        if data:
                            aggregator.add_questionnaire(data)
                            print(f"      Extracted data from {file_name}")
                        else:
                            print(f"      Failed to extract data from {file_name}")
                    else:
                        print(f"      Failed to download {file_name}")
        
        # Export aggregated data
        print("\n" + "=" * 50)
        aggregator.export_to_excel(output_file)
        print(f"\nProcessing complete! Total questionnaires processed: {len(aggregator.all_data)}")
    
    else:
        print(f"Error: Unknown mode '{mode}'. Valid modes are 'local' or 'sharepoint'")
        sys.exit(1)


if __name__ == "__main__":
    main()
