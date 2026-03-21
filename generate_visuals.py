
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import shutil
import os
import textwrap

output_dir = '/Users/sumesh/Projects/Antigravity/Capstone/presentation_visuals'
os.makedirs(output_dir, exist_ok=True)

# -----------------
# FIGURE 1: Process Overview
# -----------------
fig, ax = plt.subplots(figsize=(10, 4))
ax.axis('off')
boxes = ['Raw Vendor Data\nIngestion', 'Data Cleaning &\nStandardization', 'Filter by Sector &\nService Codes', 'Integrate Diversity\n& Perf Metrics', 'Dashboard\nVisualization']
for i, text in enumerate(boxes):
    ax.text(0.1 + i*0.2, 0.5, text, ha='center', va='center', 
            bbox=dict(boxstyle='round,pad=1', facecolor='#00BCD4', edgecolor='white', alpha=0.9), 
            fontsize=10, fontweight='bold', color='white')
    if i < len(boxes)-1:
        ax.annotate('', xy=(0.1 + i*0.2 + 0.1, 0.5), xytext=(0.1 + i*0.2 + 0.05, 0.5), 
                    arrowprops=dict(arrowstyle="->", lw=2, color='#333333'))
plt.title('Figure 1: Vendor Prioritization Process Overview', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig(f'{output_dir}/Figure1_Process_Overview.jpg', format='jpg', dpi=300)
plt.close()

# Load Data
df = pd.read_excel('/Users/sumesh/Downloads/Copy of Vendor Contact Details (1).xlsx', sheet_name=None)
SKIP_SHEETS = ['Abbreviations ']
frames = []
for sheet, data in df.items():
    if sheet.strip() in [s.strip() for s in SKIP_SHEETS]: continue
    tdf = data.iloc[:, :7].copy()
    tdf.columns = ['Contract_Code', 'Name', 'Company', 'Role', 'Email', 'Phone', 'SDO_Commitment_Pct']
    tdf['Category'] = sheet.strip()
    
    metadata_keywords = ['Master Contract', 'Solicitation Enabled', 'Master MBPO', 'Bid and Contract', 'Category and Vendor', 'Category Development', 'Mass Gov', 'OSD Help Desk', 'N/A', 'Company', 'Description']
    mask = tdf['Company'].apply(lambda x: not any(kw.lower() in str(x).lower() for kw in metadata_keywords))
    tdf = tdf[mask].copy()
    tdf['SDO_Commitment_Pct'] = pd.to_numeric(tdf['SDO_Commitment_Pct'], errors='coerce')
    tdf['Company'] = tdf['Company'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
    frames.append(tdf)

combined = pd.concat(frames, ignore_index=True)

# -----------------
# FIGURE 2: Distribution of Companies by Category
# -----------------
fig, ax = plt.subplots(figsize=(12, 8))
cat_counts = combined.groupby('Category')['Company'].nunique().sort_values(ascending=True)
bars = ax.barh(cat_counts.index, cat_counts.values, color='#7C4DFF', edgecolor='white')
ax.set_xlabel('Number of Unique Companies', fontweight='bold')
ax.set_title('Figure 2: Distribution of Companies by Category', pad=15, fontsize=14, fontweight='bold')
for bar, val in zip(bars, cat_counts.values):
    ax.text(bar.get_width() + 2, bar.get_y() + bar.get_height()/2, str(val), va='center', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/Figure2_Category_Distribution.jpg', format='jpg', dpi=300)
plt.close()

# -----------------
# FIGURE 3: Contact Verification Completeness
# -----------------
fig, ax = plt.subplots(figsize=(8, 8))
combined['Email_Valid'] = combined['Email'].apply(lambda x: 0 if pd.isna(x) or str(x).lower()=='nan' else 1)
combined['Phone_Valid'] = combined['Phone'].apply(lambda x: 0 if pd.isna(x) or str(x).lower()=='nan' else 1)
combined['Contact_Verified'] = combined.apply(lambda row: 1 if row['Email_Valid']==1 and row['Phone_Valid']==1 else 0, axis=1)

verified = combined.groupby('Company')['Contact_Verified'].max()
counts = [verified.sum(), len(verified) - verified.sum()]
labels = ['Verified Contact', 'Missing/Incomplete Contact']
colors = ['#00BCD4', '#FF6F61']

wedges, texts, autotexts = ax.pie(counts, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors,
                                  wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2),
                                  textprops=dict(fontweight='bold', fontsize=11))
ax.set_title('Figure 3: Contact Verification Completeness', pad=20, fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/Figure3_Contact_Verification.jpg', format='jpg', dpi=300)
plt.close()

# -----------------
# FIGURE 4: Top 10 Companies by Average SDO Commitment
# -----------------
sdo_df = combined.dropna(subset=['SDO_Commitment_Pct']).copy()
sdo_df = sdo_df[(sdo_df['SDO_Commitment_Pct'] > 0) & (sdo_df['SDO_Commitment_Pct'] <= 1.0)]
top10_sdo = sdo_df.groupby('Company')['SDO_Commitment_Pct'].mean().sort_values(ascending=False).head(10)

fig, ax = plt.subplots(figsize=(12, 7))
labels_wrap = [textwrap.fill(x, 30) for x in top10_sdo.index]
bars = ax.barh(labels_wrap, top10_sdo.values * 100, color='#FF6F61', edgecolor='white')
ax.invert_yaxis()
ax.set_xlabel('Average SDO Commitment (%)', fontweight='bold')
ax.set_title('Figure 4: Top 10 Companies by Average SDO Commitment', pad=15, fontsize=14, fontweight='bold')
for bar, val in zip(bars, top10_sdo.values):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2, f"{val:.1%}", va='center', fontweight='bold')
plt.tight_layout()
plt.savefig(f'{output_dir}/Figure4_Top10_SDO.jpg', format='jpg', dpi=300)
plt.close()

# -----------------
# FIGURE 5: Interactive Plotly Dash Deployment
# -----------------
shutil.copy('/Users/sumesh/Projects/Antigravity/Capstone/dashboard_screenshot.png', f'{output_dir}/Figure5_Dashboard.jpg')

print("All figures generated successfully.")
