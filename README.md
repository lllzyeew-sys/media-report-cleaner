# PLAUD Media Report Cleaner - Web App

A deployable web tool for cleaning PLAUD media monitoring CSV reports.

## What it does

Upload a CSV report, then the app automatically:

1. Reads common CSV encodings and delimiters.
2. Detects likely fields for Title, Hit Sentence, Source, URL, and Country/Region.
3. Classifies each row as Keep / Remove / Review.
4. Removes Medium and Substack sources before other content checks.
5. Splits Keep rows by region: Americas, JP, EU, APAC excluding JP, and Other/Unassigned.
6. Exports a full Excel report plus separate CSV files.

## Deploy option 1: Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload all files in this folder to the repository.
3. Go to Streamlit Community Cloud and create a new app from the repository.
4. Set the main file path to `app.py`.
5. Deploy.

This is the easiest hosted option if company policy allows using Streamlit Cloud.

## Deploy option 2: Render

1. Create a GitHub repository.
2. Upload all files in this folder to the repository.
3. In Render, create a new Web Service from the repository.
4. Render can use the included `render.yaml` and `Dockerfile`.
5. Deploy.

## Deploy option 3: Internal server with Docker

From this folder:

```bash
docker build -t plaud-media-report-cleaner .
docker run -p 8501:8501 plaud-media-report-cleaner
```

Then open:

```text
http://SERVER_IP:8501
```

## Local run for testing

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- Uploaded CSV files are processed in memory during the session.
- The app does not require a database.
- For production/internal use, deploy behind your company SSO or VPN if the reports contain sensitive media monitoring data.
