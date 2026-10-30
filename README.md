---
title: Fiber 37 Tools
emoji: 🧵
colorFrom: blue
colorTo: green
sdk: gradio
app_file: app.py
pinned: false
license: mit
---

# Fiber OSP Cost Estimator

Combined Longhaul OSP Assessment + Market / MDU Cost Model for Hugging Face Spaces.

## Features
- **Longhaul OSP Tab**: Route miles, plow/bore footage, unit rates, rock premium, contingency, duration.
- **Market / MDU Tab**: Passings, aerial/UG split, MDU incremental costs, take rate, capex per passing, Year-1 revenue.
- Live calculation with editable inputs based on the original spreadsheet models.
- **Exports**: Excel workbooks and NotebookLM-ready Markdown files.

## Deploy to Hugging Face Spaces
1. Create a new Space (Gradio SDK) or link this repo.
2. The `app.py` and `requirements.txt` are ready.
3. Push. The app will launch automatically.

## Local Run
```bash
pip install -r requirements.txt
python app.py
```

All rates are preliminary planning values. Validate against current vendor quotes and local conditions before using for decisions.
