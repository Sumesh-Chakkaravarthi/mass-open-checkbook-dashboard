# Massachusetts Open Checkbook Dashboard 📊

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://www.python.org/)
[![Dash](https://img.shields.io/badge/Dash-Plotly-informational.svg)](https://dash.plotly.com/)
[![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-red.svg)](https://pandas.pydata.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**[🔴 LIVE DEMO: Click Here](https://sumesh-chakkaravarthi.github.io/mass-open-checkbook-dashboard/)**

A comprehensive, self-contained data project exploring the **Massachusetts Open Checkbook** vendor contract data. This repository features an end-to-end data pipeline from Exploratory Data Analysis (EDA) to a dynamic, interactive dashboard built with Plotly Dash.

![Dashboard Overview](presentation_visuals/Figure5_Dashboard.jpg)

---

## 📖 Context & Background
The **Supplier Diversity Office (SDO)** of Massachusetts promotes diversity, equity, and inclusion in state contracting. This project analyzes vendor contracts across the IT Sector, identifying disparities in SDO commitment and exploring geographical distribution. 

### Key Business Questions Analyzed:
- What is the vendor concentration in the IT sector?
- How do National vs. Local SDO performances compare?
- Are certain contract subcategories dominated by specific vendors?
- How is the structural density of SDO companies distributed?

## 📂 Dataset
- **Source**: Massachusetts Open Checkbook Public Data
- **Focus Area**: IT Sector, Cross-Category Analysis, and Vendor Coverage.
- Data files are processed via `eda_analysis.py` to generate the interactive visualizations and analytical models.

## 📈 Key Findings & Visualizations

### 1. Vendor Distribution
![Vendor Distribution Treemap](output/BQ3_Vendor_Distribution_Treemap.png)
*A high-level view of vendor category distribution, indicating areas with the highest contract concentration.*

### 2. SDO Companies in IT
![Top SDO Companies in IT](output/BQ1_IT_Top_SDO_Companies.png)
*Highlighting the top performing Supplier Diversity Office certified companies within the IT sector and their project volume.*

### 3. Industry Diversity 
![Industry Diversity Heatmap](output/BQ9_Industry_Diversity_Heatmap.png)
*A visual heatmap showcasing the correlation between diverse industries and SDO coverage rates across multiple dimensions.*

---

## 🚀 How to Run Locally

1. **Clone the repository**:
   ```bash
   git clone https://github.com/Sumesh-Chakkaravarthi/mass-open-checkbook-dashboard.git
   cd mass-open-checkbook-dashboard
   ```

2. **Create a virtual environment**:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Launch the Dashboard**:
   ```bash
   python dashboard.py
   ```
   Open your browser and navigate to `http://127.0.0.1:8050` to view the interactive application.

5. *(Optional) Run the EDA script to regenerate visualizations:*
   ```bash
   python eda_analysis.py
   ```

## 🧠 About
This project was developed as a Capstone to demonstrate proficiency in:
- **Programming**: Python
- **Data Manipulation**: Pandas, NumPy
- **Data Visualization**: Plotly, Dash, Matplotlib, Seaborn
- **Interactive Apps**: Dash Web Applications

Authored by **Sumesh Chakkaravarthi**.
