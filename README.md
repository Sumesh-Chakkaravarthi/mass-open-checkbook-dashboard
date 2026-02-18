# Massachusetts Open Checkbook Dashboard

**[🔴 LIVE DEMO: Click Here](https://sumesh-chakkaravarthi.github.io/mass-open-checkbook-dashboard/)**

This is a **Plotly Dash** application designed to analyze vendor contracts and SDO (Supplier Diversity Office) commitments for the Commonwealth of Massachusetts.

## Features
- **Interactive Dashboard**: 4 tabs covering IT Sector, Cross-Category analysis, Vendor Coverage, and Industry Analysis.
- **Data Visualization**: Over 10 interactive charts including bar charts, treemaps, scatter plots, and radar charts.
- **Search & Filter**: Filter IT companies by sub-category (Hardware, Software, Telecom).

## Installation (For Developers)
1. Clone the repository:
   ```bash
   git clone https://github.com/Sumesh-Chakkaravarthi/mass-open-checkbook-dashboard.git
   cd mass-open-checkbook-dashboard
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Run the dashboard locally:
```bash
python dashboard.py
```
Open your browser to `http://127.0.0.1:8050`.

## Data Sources
- **Vendor Data**: `Copy of Vendor Contact Details (1).xlsx`
- **Categorized Companies**: `List of Categorized_Companies (1).xlsx`
