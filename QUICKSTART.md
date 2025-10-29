# Quick Start Guide

This guide will help you get started with the Folder Aggregator.

## Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

## Step 2: Configure Your Settings

1. Copy the example environment file:
```bash
cp .env.example .env
```

2. Edit `.env` with your settings:

### Local Mode (Default - Recommended)

```env
MODE=local
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

This will scan all folders in the current directory (where you run the script from).

### For SharePoint Mode

```env
MODE=sharepoint
SHAREPOINT_SITE_URL=https://yourcompany.sharepoint.com/sites/yoursite
SHAREPOINT_USERNAME=your.email@company.com
SHAREPOINT_PASSWORD=your_password
QUESTIONNAIRE_FOLDER_PATH=Shared Documents/questionnaire
OUTPUT_FILE=aggregated_questionnaires.xlsx
```

### Finding Your SharePoint Site URL (SharePoint Mode Only)

1. Open your SharePoint site in a web browser
2. Copy the URL up to `/sites/yoursite` (before any document paths)
3. Example: `https://contoso.sharepoint.com/sites/HR`

### Finding Your Folder Path (SharePoint Mode Only)

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

### Local Mode

The tool will:
1. Scan all folders in the current working directory
2. Find Excel files in each folder
3. Extract and aggregate the data
4. Create a single Excel file with all results

**Example:**
If your folder structure looks like this:
```
/your/directory/
├── sharepoint_aggregator.py
├── Department1/
│   └── questionnaire.xlsx
├── Department2/
│   └── questionnaire.xlsx
└── Department3/
    └── questionnaire.xlsx
```

The script will process all Excel files from Department1, Department2, and Department3.

### SharePoint Mode

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

### Local Mode

**No folders found:**
- Make sure there are folders in the same directory as the script
- Folders starting with `.` (hidden folders) are ignored

**No Excel files found:**
- Ensure files have `.xlsx` or `.xls` extensions
- Files starting with `~$` (temporary Excel files) are ignored

### SharePoint Mode

#### Authentication Fails

If you have Multi-Factor Authentication (MFA) enabled:
1. Go to https://account.microsoft.com/security
2. Create an app password
3. Use the app password in your `.env` file instead of your regular password

#### Cannot Find Folder

Try different path formats:
- `/sites/yoursite/Shared Documents/questionnaire`
- `Shared Documents/questionnaire`
- Check the exact folder name in SharePoint (case-sensitive)

#### Permission Denied

Ensure your account has:
- Read access to the SharePoint folder
- Permission to list folder contents
- Permission to download files

Contact your SharePoint administrator if needed.

## Need Help?

Check the full README.md for detailed documentation and troubleshooting tips.
