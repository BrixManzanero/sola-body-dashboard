# Sola Body — Revenue Command Center

## First time setup (do this once)

Open the VS Code terminal (Ctrl + ` ) and run:

```
pip install -r requirements.txt
```

## Run the dashboard

**Option 1 — Press F5** (VS Code will launch it automatically)

**Option 2 — Terminal:**
```
python app.py
```

Then open your browser at **http://localhost:5000**

---

## How to use

1. Upload your **BR Input Excel** file (drag & drop on the page)
2. Select a date from the sidebar
3. View your revenue, ad spend, ROAS, and charts
4. Click any **Download** button for the Excel report you need

---

## What you need to upload

| | |
|---|---|
| File | Your daily BR Input Excel (e.g. `2026_05_25_BR_Input.xlsx`) |
| Required sheet | `Daily Revenue and Ads Expense!` |
| Columns needed | Shopify, Shopee, Lazada, Tiktok Shop, all ad spend columns |

---

## Files in this project

| File | What it does |
|---|---|
| `app.py` | Python backend — reads Excel, cleans data, generates reports |
| `templates/index.html` | Dashboard page — HTML structure only |
| `static/css/style.css` | All styling — colors, layout, cards, charts |
| `static/js/main.js` | All interactivity — file upload, charts, API calls |
| `requirements.txt` | Python packages needed |
| `.vscode/launch.json` | Lets you run with F5 in VS Code |
