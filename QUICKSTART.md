# Quick Start Guide

This guide will help you get started with the SharePoint Questionnaire Aggregator.

## Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

## Step 2: Configure Your Settings

1. Copy the example environment file:
```bash
cp .env.example .env
```

2. Edit `.env` with your details:
```env
SHAREPOINT_SITE_URL=https://yourcompany.sharepoint.com/sites/yoursite
SHAREPOINT_USERNAME=your.email@company.com
SHAREPOINT_PASSWORD=your_password
QUESTIONNAIRE_FOLDER_PATH=Shared Documents/questionnaire
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

### Finding Your SharePoint Site URL

1. Open your SharePoint site in a web browser
2. Copy the URL up to `/sites/yoursite` (before any document paths)
3. Example: `https://contoso.sharepoint.com/sites/HR`

### Finding Your Folder Path

The folder path should be relative to the site. Common formats:
- `Shared Documents/questionnaire`
- `Documents/questionnaire`

To find the exact path:
1. Navigate to your questionnaire folder in SharePoint
2. Look at the URL - it will contain the path after the site URL
3. Use the part after the site URL

## Step 3: Prepare Your Excel Files

Ensure your questionnaire Excel files follow this structure:

- **Row 1**:
  - Cell B1: Application Name
  - Cell C1: Responsible Person
  - Cell D1: Deputy Person

- **Rows 3-19** (17 questions):
  - Column B: Question text
  - Column C: Primary answer
  - Column D: Alternative answer (used if Column C is empty or contains "None")

See `example_questionnaire.xlsx` for a template.

## Step 4: Run the Tool

```bash
python sharepoint_aggregator.py
```

The tool will:
1. Connect to SharePoint
2. Scan all folders in the questionnaire directory
3. Download Excel files from each folder
4. Extract and aggregate the data
5. Create a single Excel file with all results

## Step 5: Review the Output

Open the generated Excel file (default: `aggregated_questionnaires.xlsx`).

It will contain:
- One row per questionnaire
- Columns: Source Folder, Application, Responsible, Deputy, Question 1, Answer 1, ..., Question 17, Answer 17

## Troubleshooting

### Authentication Fails

If you have Multi-Factor Authentication (MFA) enabled:
1. Go to https://account.microsoft.com/security
2. Create an app password
3. Use the app password in your `.env` file instead of your regular password

### Cannot Find Folder

Try different path formats:
- `/sites/yoursite/Shared Documents/questionnaire`
- `Shared Documents/questionnaire`
- Check the exact folder name in SharePoint (case-sensitive)

### Permission Denied

Ensure your account has:
- Read access to the SharePoint folder
- Permission to list folder contents
- Permission to download files

Contact your SharePoint administrator if needed.

## Need Help?

Check the full README.md for detailed documentation and troubleshooting tips.
