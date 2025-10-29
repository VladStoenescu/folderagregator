# SharePoint Questionnaire Aggregator

A Python tool to automatically download Excel questionnaires from SharePoint folders and aggregate the data into a single Excel file.

## Features

- Connects to SharePoint using Office 365 authentication
- Navigates through all subfolders within a specified SharePoint folder
- Downloads Excel files from each subfolder
- Extracts questionnaire data including:
  - Application name (Cell B1)
  - Responsible person (Cell C1)
  - Deputy person (Cell D1)
  - 17 questions with answers (Rows 3-19, Columns B for questions, C & D for answers)
- Aggregates all data into a single Excel file with organized columns

## Prerequisites

- Python 3.8 or higher
- SharePoint Online account with appropriate permissions
- Access to the SharePoint site and folder containing questionnaires

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

2. Edit the `.env` file with your SharePoint credentials and settings:

```env
# SharePoint Configuration
SHAREPOINT_SITE_URL=https://yourcompany.sharepoint.com/sites/yoursite
SHAREPOINT_USERNAME=your.email@company.com
SHAREPOINT_PASSWORD=your_password
QUESTIONNAIRE_FOLDER_PATH=Shared Documents/questionnaire
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

### Configuration Parameters:

- **SHAREPOINT_SITE_URL**: Full URL to your SharePoint site
  - Example: `https://contoso.sharepoint.com/sites/HR`
  
- **SHAREPOINT_USERNAME**: Your SharePoint/Microsoft 365 email address
  
- **SHAREPOINT_PASSWORD**: Your SharePoint/Microsoft 365 password
  - **Note**: For accounts with Multi-Factor Authentication (MFA), you may need to use an app password
  
- **QUESTIONNAIRE_FOLDER_PATH**: Relative path to the questionnaire folder
  - Example: `Shared Documents/questionnaire` or `Documents/questionnaire`
  
- **OUTPUT_FILE**: Name of the output Excel file (optional, defaults to `aggregated_questionnaires.xlsx`)

## Usage

Run the script:
```bash
python sharepoint_aggregator.py
```

The script will:
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
  - Cell C1: Responsible Person
  - Cell D1: Deputy Person

- **Rows 3-19**:
  - Column B: Questions (17 questions total)
  - Column C: Primary answers
  - Column D: Alternative answers (used if Column C is empty or contains "None")

## Output Format

The aggregated Excel file will contain columns:
- Source Folder
- Application
- Responsible
- Deputy
- Question 1, Answer 1
- Question 2, Answer 2
- ... (up to Question 17, Answer 17)

Each row represents one questionnaire from the SharePoint folders.

## Troubleshooting

### Authentication Issues

If you encounter authentication errors:

1. **Multi-Factor Authentication (MFA)**: If your account uses MFA, you need to create an app password:
   - Go to Microsoft account security settings
   - Create a new app password
   - Use this app password in your `.env` file instead of your regular password

2. **Modern Authentication**: Some organizations require modern authentication. In this case, you may need to:
   - Contact your IT administrator to enable legacy authentication for your app
   - Or use alternative authentication methods (Azure AD app registration)

### Permission Issues

Ensure your SharePoint account has:
- Read access to the questionnaire folder and all subfolders
- Ability to list folder contents
- Ability to download files

### Path Issues

- Make sure the `QUESTIONNAIRE_FOLDER_PATH` is correct
- Try using the server-relative path format: `/sites/yoursite/Shared Documents/questionnaire`
- Or use the simplified format: `Shared Documents/questionnaire`

### Connection Issues

- Verify the SharePoint site URL is correct
- Check if you can access the site through a web browser
- Ensure your network allows connections to SharePoint

## Security Notes

- **Never commit your `.env` file** to version control (it's already in `.gitignore`)
- Store credentials securely
- Consider using Azure AD app registration for production environments
- Rotate passwords regularly

## Dependencies

- `Office365-REST-Python-Client`: SharePoint REST API client
- `openpyxl`: Excel file reading and writing
- `pandas`: Data manipulation and Excel export
- `python-dotenv`: Environment variable management

## License

This project is provided as-is for internal use.

## Support

For issues or questions, please create an issue in the repository.
