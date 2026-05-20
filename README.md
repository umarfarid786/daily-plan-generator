# Daily Team Plan Generator

This is a beginner-friendly Streamlit app for generating daily team-wise inspection plans from an Excel workbook.

## What it does

- Upload an Excel file
- Select the sheet containing the master inspection plan
- Select Day No and Team No
- Generate the filtered daily plan
- Download one team plan as Excel
- Download all teams for one day as a ZIP file

## Required columns

The app tries to detect these columns automatically:

- Day No
- Team No
- Visit Sequence
- Team Day Town
- Business Name
- Address
- FBO Name
- Contact
- Planned Category

If a column is not detected, the app will ask you to map it manually.

## Local setup

1. Install Python from https://www.python.org/downloads/
2. Open Command Prompt or Terminal in this folder.
3. Run:

```bash
pip install -r requirements.txt
```

4. Start the app:

```bash
streamlit run app.py
```

5. Open the local link shown in the terminal, usually:

```text
http://localhost:8501
```

## Online deployment later

After testing locally, upload these files to GitHub and deploy through Streamlit Community Cloud:

- app.py
- requirements.txt
- README.md

Do not upload confidential master Excel files to a public GitHub repository.
