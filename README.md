# Attendance Dashboard — Streamlit

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then use the file picker in the sidebar to upload an `.xlsx`, `.xls`, or `.csv` attendance file.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py` and `requirements.txt`.
3. In Streamlit Community Cloud, create a new app and select `app.py` as the main file.
4. Deploy.

No local file paths, Tkinter, or OS-specific dialogs are used. The Streamlit uploader opens the browser's native file-selection dialog, so it works on Streamlit Cloud.

## Supported Excel register

The parser is designed for the supplied attendance-register layout:
- metadata rows for department, academic year, semester, year, division and reporting period
- student rows identified by numeric Roll No.
- three-column subject groups using `Pre / Per`
- final `Overall TH Att.`, `Overall PR Att.`, and `Overall Att.` columns
- bottom `Average` and `Name of Faculty` rows

CSV support accepts conventional columns such as:
`Roll No`, `Name`, subject attendance columns, and optional overall attendance columns.

## Dashboard sections

- Class Snapshot
- Attendance distribution
- Subject-wise average attendance
- Key insights
- Top 5 / Bottom 5
- Search and attendance filters
- Student detail by subject
- Faculty & Subject Summary
