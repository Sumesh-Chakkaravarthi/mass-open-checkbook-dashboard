# Massachusetts Open Checkbook Dashboard

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Dash](https://img.shields.io/badge/Dash-Plotly-0D6EFD.svg)](https://dash.plotly.com/)
[![Stack](https://img.shields.io/badge/Stack-Pandas_|_Scikit--Learn-150458.svg)](https://scikit-learn.org/)
[![CI](https://github.com/Sumesh-Chakkaravarthi/mass-open-checkbook-dashboard/actions/workflows/ci.yml/badge.svg)](https://github.com/Sumesh-Chakkaravarthi/mass-open-checkbook-dashboard/actions)

**[Live Demo](https://sumesh-chakkaravarthi.github.io/mass-open-checkbook-dashboard/)**

---

## Overview

An end-to-end data analytics project analyzing **Massachusetts Open Checkbook** vendor contract data. The pipeline covers data ingestion, exploratory analysis, predictive modeling, and interactive visualization — focused on Supplier Diversity Office (SDO) commitment metrics across 15 procurement categories and 10 industry classifications.

---

## Architecture

```
Data Sources (Excel)
    |
    v
EDA & Feature Engineering (eda_analysis.py)
    |
    v
ML Pipeline (train_sdo_model.py)
    |   - OneHotEncoder + ColumnTransformer
    |   - Random Forest Regressor
    |   - Feature importance analysis
    v
Interactive Dashboard (dashboard.py)
    |   - 4-tab Plotly Dash application
    |   - KPI cards, filters, drill-downs
    v
Static Export (static_builder.py -> index.html)
    |   - GitHub Pages deployment
```

---

## Key Features

- **Predictive Modeling** — Random Forest Regressor predicts vendor SDO commitment based on categorical features (contract category, vendor role). Feature importance analysis identifies the strongest predictors.
- **Interactive Dashboard** — Multi-tab Dash application with real-time filtering, KPI summary cards, and 10+ visualizations across IT sector analysis, cross-category comparison, vendor coverage, and industry diversity.
- **Automated CI/CD** — GitHub Actions pipeline validates code quality and dependency integrity on every push.

---

## Machine Learning Pipeline

The `train_sdo_model.py` script preprocesses categorical vendor features using `OneHotEncoder`, builds a `ColumnTransformer` pipeline, and trains a `RandomForestRegressor` to predict SDO commitment percentages.

### Feature Importance
![Feature Importance](output/ML_Feature_Importance.png)

---

## Visualizations

### Vendor Distribution Across Categories
![Vendor Distribution](output/BQ3_Vendor_Distribution_Treemap.png)

### Top IT Sector Companies by SDO Commitment
![Top SDO Companies](output/BQ1_IT_Top_SDO_Companies.png)

### Industry Diversity Heatmap
![Industry Diversity](output/BQ9_Industry_Diversity_Heatmap.png)

---

## Project Structure

```
.
├── dashboard.py            # Interactive Plotly Dash application
├── eda_analysis.py         # Exploratory data analysis (15 business questions)
├── train_sdo_model.py      # ML training pipeline
├── static_builder.py       # Generates static HTML for GitHub Pages
├── generate_ppt.py         # PowerPoint report generator
├── generate_report.py      # Word document report generator
├── generate_visuals.py     # Presentation visual generator
├── index.html              # Static dashboard (GitHub Pages)
├── requirements.txt        # Python dependencies
├── models/
│   └── sdo_rf_model.pkl    # Trained Random Forest model
├── output/                 # Generated EDA visualizations (17 PNGs)
├── presentation_visuals/   # High-level presentation figures
└── .github/workflows/
    └── ci.yml              # CI pipeline configuration
```

---

## Getting Started

```bash
# Clone
git clone https://github.com/Sumesh-Chakkaravarthi/mass-open-checkbook-dashboard.git
cd mass-open-checkbook-dashboard

# Environment
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Launch dashboard
python dashboard.py
# Open http://127.0.0.1:8050
```

---

## Tech Stack

| Layer | Tools |
|-------|-------|
| Data Processing | Pandas, NumPy, OpenPyXL |
| Machine Learning | Scikit-Learn (Random Forest, OneHotEncoder) |
| Visualization | Plotly, Matplotlib, Seaborn |
| Web Framework | Plotly Dash |
| CI/CD | GitHub Actions |
| Deployment | GitHub Pages |

---

## Author

**Sumesh Chakkaravarthi**
