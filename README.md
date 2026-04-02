# Rent Roll AI Agent

A true AI agent that auto-detects any multi-family rent roll format, extracts the right data, and populates your standard Rent Roll template — powered by Claude.

## What it does

- **Auto-detects** the source system format (Yardi, RealPage/OneSite, MRI, AppFolio, Rent Manager, and more)
- **Extracts** unit data, resident info, dates, market rents, and effective rents
- **Handles** multi-row charge codes — sums only true rent codes, ignores fees
- **Skips** applicants, future residents, and pending renewals
- **Zeros** effective rent for vacant, model, and down units
- **Groups** units by building/section for unit type
- **Self-validates** output against source totals
- **Loops** with Claude AI correction if validation fails

## Supported formats (built-in parsers)

| Parser | Systems |
|--------|---------|
| Yardi | Yardi Voyager, Yardi Breeze |
| OneSite/RealPage | OneSite Rents v3.0 |
| MRI Living | MRI Living (charge-code and actual-rent variants) |
| AppFolio | AppFolio, similar multi-section formats |
| Rent Manager | Rent Manager, WAT-style exports |
| **AI Fallback** | Any other system via Claude schema detection |

## Deploying on Streamlit Cloud (no installation needed)

### Step 1 — Create a GitHub repo

1. Go to [github.com](https://github.com) and create a **new repository** (e.g. `rent-roll-agent`)
2. Upload these three files:
   - `app.py`
   - `requirements.txt`
   - `Rent_Roll_template.xlsx`

### Step 2 — Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
2. Click **"New app"**
3. Select your repository, branch (`main`), and set **Main file path** to `app.py`
4. Click **"Deploy"**

That's it! Streamlit Cloud installs all dependencies automatically.

### Step 3 — Set your Anthropic API key (for AI fallback)

The app uses your Anthropic API key for schema detection on unknown formats.

In Streamlit Cloud:
1. Go to your app settings → **Secrets**
2. Add:
```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

> **Note:** The 5 built-in parsers (Yardi, OneSite, MRI, AppFolio, Rent Manager) work without an API key. The key is only used for the AI fallback on unknown formats.

## How to use

1. Open the deployed app URL
2. Set the **Rent Roll As-Of Date** in the sidebar
3. Upload your raw rent roll `.xlsx` file
4. Click **"Run Agent"**
5. Review the results and download the filled template

## Adding the API key to the app

The app reads the key from Streamlit secrets automatically. If running locally, create `.streamlit/secrets.toml`:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

## Running locally (optional)

```bash
pip install streamlit openpyxl pandas requests
streamlit run app.py
```

## File requirements

- **Input:** Any `.xlsx` rent roll export from a property management system
- **Output:** Filled `Rent_Roll_template.xlsx` with your standard columns populated
- The template's existing formulas and formatting are preserved

## Accuracy on test files

| File | System | Accuracy |
|------|--------|----------|
| Nova Central | Yardi Voyager | 100% |
| One Hampton Lake | AppFolio | 100% |
| Stone Loch | MRI Living | 100% |
| WAT Waterford Place | Rent Manager | 100% |
| Retreat at the Park | MRI Living | 100% |
| The Retreat at Canal | MRI Living | 100% |
| Timberlakes | RealPage OneSite | 97.4% |
| ReNew Western Cranston | AppFolio | ~95% |
