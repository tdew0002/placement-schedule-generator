# Placement Schedule Generator

A Streamlit web app that reads a Google Form CSV export of student placement availability and generates one formatted Excel file per school.

## How to use

1. Open the app URL
2. Upload the Google Form CSV export
3. Choose which schools to export
4. Click Generate and download the ZIP

## Local development

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
