# Pickleball Player Manager

A responsive Streamlit tournament dashboard for Google Forms registration uploads, DUPR lookups, player classification, filtering, and formatted Excel exports.

## Run locally

1. Install Python 3.12 or another version supported by Streamlit Community Cloud.
2. Install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

3. Create `.streamlit/secrets.toml`:

   ```toml
   DUPR_TOKEN = "YOUR_REAL_DUPR_TOKEN"
   ```

4. Start the app:

   ```powershell
   streamlit run Main.py
   ```

On Windows, `Run-Pickleball.ps1` performs the token setup and starts the app.

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repository. Keep the repository root unchanged so `Main.py` and `requirements.txt` remain at the top level.
2. In Streamlit Community Cloud, create an app and select the repository, branch, and `Main.py` entry point.
3. Open the app's advanced settings and add this secret:

   ```toml
   DUPR_TOKEN = "YOUR_REAL_DUPR_TOKEN"
   ```

4. Deploy the app. Do not commit `.streamlit/secrets.toml`; it is excluded by `.gitignore`.

The interface uses responsive wrapping, horizontally scrollable navigation, touch-sized controls, and horizontally scrollable data tables for phones and tablets.
