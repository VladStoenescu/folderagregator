# Folder Aggregator

A Python tool to automatically process Excel questionnaires from local folders or SharePoint, and aggregate the data into a single Excel file.

## Features

- **Local Mode (Default)**: Scans all folders in the current working directory (where you run the script)
- **SharePoint Mode**: Connects to SharePoint and navigates through subfolders
- Downloads/reads Excel files from each folder
- Extracts questionnaire data including:
  - Application name (Cell B1)
  - Responsible person (Cell C1)
  - Deputy person (Cell D1)
  - 17 questions with answers (Rows 3-19, Columns B for questions, C & D for answers)
- Aggregates all data into a single Excel file with organized columns

## Prerequisites

- Python 3.8 or higher
- For SharePoint mode: SharePoint Online account with appropriate permissions

## Installation

1. Clone this repository or download the files

2. Install required Python packages:
```bash
pip install -r requirements.txt
```

## Configuration

1. Copy the `.env.example` file to `.env`:
```bash
cp .env.example .env
```

2. Edit the `.env` file with your settings:

### Local Mode (Default)

```env
MODE=local
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

The script will scan all folders in the current working directory (where you run the script from).

Optionally, you can specify a different base path:
```env
MODE=local
BASE_PATH=/path/to/your/folders
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

### SharePoint Mode

To use SharePoint mode, set `MODE=sharepoint` and provide SharePoint credentials:

```env
MODE=sharepoint
SHAREPOINT_SITE_URL=https://yourcompany.sharepoint.com/sites/yoursite
SHAREPOINT_USERNAME=your.email@company.com
SHAREPOINT_PASSWORD=your_password
QUESTIONNAIRE_FOLDER_PATH=Shared Documents/questionnaire
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

### Configuration Parameters:

**Common Parameters:**
- **MODE**: Operating mode - `local` (default) or `sharepoint`
- **OUTPUT_FILE**: Name of the output Excel file (optional, defaults to `aggregated_questionnaires.xlsx`)

**Local Mode Parameters:**
- **BASE_PATH**: Directory to scan for folders (optional, defaults to current working directory)

**SharePoint Mode Parameters:**
- **SHAREPOINT_SITE_URL**: Full URL to your SharePoint site
  - Example: `https://contoso.sharepoint.com/sites/HR`
  
- **SHAREPOINT_USERNAME**: Your SharePoint/Microsoft 365 email address
  
- **SHAREPOINT_PASSWORD**: Your SharePoint/Microsoft 365 password
  - **Note**: For accounts with Multi-Factor Authentication (MFA), you may need to use an app password
  
- **QUESTIONNAIRE_FOLDER_PATH**: Relative path to the questionnaire folder
  - Example: `Shared Documents/questionnaire` or `Documents/questionnaire`

## Usage

Run the script:
```bash
python sharepoint_aggregator.py
```

### Local Mode (Default)

The script will:
1. Scan all folders in the same directory as the script
2. Look for Excel files in each folder
3. Extract the questionnaire data from each Excel file
4. Create an aggregated Excel file with all data

**Example folder structure:**
```
/your/directory/
├── sharepoint_aggregator.py
├── Folder1/
│   ├── questionnaire1.xlsx
│   └── questionnaire2.xlsx
├── Folder2/
│   └── questionnaire3.xlsx
└── Folder3/
    └── questionnaire4.xlsx
```

The script will process all Excel files in Folder1, Folder2, and Folder3.

### SharePoint Mode

When configured for SharePoint mode, the script will:
1. Connect to SharePoint using your credentials
2. Navigate to the specified questionnaire folder
3. Scan all subfolders for Excel files
4. Download and process each Excel file
5. Extract the questionnaire data
6. Create an aggregated Excel file with all data

## Expected Excel File Structure

The tool expects Excel files with the following structure:

- **Row 1**:
  - Cell B1: Application Name
  - Cell C1: Responsible Person (Owner)
  - Cell D1: Deputy Person

- **Rows 3-19**:
  - Column A: Serial Number (1-17)
  - Column B: Questions (17 questions total)
  - Column C: Answers
  - Column D: Comments (optional)

## Output Format

The aggregated Excel file will contain columns:
- **Application**: The application name
- **Answered**: The responsible person (owner)
- **App Responsible**: The responsible person (owner, duplicated for compatibility)
- **Deputy**: The deputy person
- **Q1 to Q17**: Answers to questions 1 through 17
- **COMM Q1 to COMM Q17**: Comments for questions 1 through 17

Each row represents one questionnaire from the folders.

## Troubleshooting

### Local Mode Issues

**No folders found:**
- Ensure there are subfolders in the current working directory
- Check that folder names don't start with a dot (hidden folders are ignored)
- Verify the BASE_PATH in .env if you specified a custom path
- You can specify BASE_PATH=/absolute/path/to/folders in .env to scan a different location

**No Excel files found:**
- Ensure Excel files have `.xlsx` or `.xls` extensions
- Temporary Excel files starting with `~$` are automatically ignored

### SharePoint Mode Issues

#### Authentication Issues

If you encounter authentication errors:

1. **Multi-Factor Authentication (MFA)**: If your account uses MFA, you need to create an app password:
   - Go to Microsoft account security settings
   - Create a new app password
   - Use this app password in your `.env` file instead of your regular password

2. **Modern Authentication**: Some organizations require modern authentication. In this case, you may need to:
   - Contact your IT administrator to enable legacy authentication for your app
   - Or use alternative authentication methods (Azure AD app registration)

#### Permission Issues

For SharePoint mode, ensure your account has:
- Read access to the questionnaire folder and all subfolders
- Ability to list folder contents
- Ability to download files

#### Path Issues

- Make sure the `QUESTIONNAIRE_FOLDER_PATH` is correct
- Try using the server-relative path format: `/sites/yoursite/Shared Documents/questionnaire`
- Or use the simplified format: `Shared Documents/questionnaire`

#### Connection Issues

- Verify the SharePoint site URL is correct
- Check if you can access the site through a web browser
- Ensure your network allows connections to SharePoint

## Security Notes

- **Never commit your `.env` file** to version control (it's already in `.gitignore`)
- Store credentials securely
- Consider using Azure AD app registration for production environments
- Rotate passwords regularly

## Dependencies

- `openpyxl`: Excel file reading and writing
- `pandas`: Data manipulation and Excel export
- `python-dotenv`: Environment variable management
- `Office365-REST-Python-Client`: SharePoint REST API client (optional, only needed for SharePoint mode)

## License

This project is provided as-is for internal use.

## Support

For issues or questions, please create an issue in the repository.
