#!/usr/bin/env python3
"""
ALY 6980 Capstone - Exploratory Data Analysis (EDA)
Massachusetts Open Checkbook: Vendor Contract & SDO Commitment Analysis

This script performs comprehensive EDA on Massachusetts state procurement
vendor contact data, analyzing SDO (Supplier Diversity Office) commitment
percentages across 15 procurement categories and 10 industry classifications.

Business Questions Analyzed:
  BQ1: Top IT sector companies by SDO commitment %
  BQ2: Average SDO commitment % by procurement category
  BQ3: Distribution of vendors across procurement categories (treemap)
  BQ4: Vendor count by contract sub-category codes
  BQ5: National vs Local company presence across industries
  BQ6: SDO commitment distribution & outlier detection
  BQ7: SDO coverage rate — vendors with valid SDO vs missing
  BQ8: Correlation between vendor count and avg SDO by category
  BQ9: Industry diversity heatmap (National vs Local vs SGC Target)
  BQ10: Vendor concentration in IT sector (market dominance)

Author: Sumesh Chakkaravarthi
Date: February 2026
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from pathlib import Path
import warnings
import textwrap

warnings.filterwarnings('ignore')

# ──────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────
VENDOR_FILE = "Copy of Vendor Contact Details (1).xlsx"
CATEGORIZED_FILE = "List of Categorized_Companies (1).xlsx"
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# Category abbreviation mapping
CATEGORY_NAMES = {
    'ITE': 'IT Equipment & Services',
    'ITS': 'IT Software & Services',
    'ITT': 'Telecom & Networking',
    'FAC': 'Facilities General',
    'VEH': 'Vehicle Acquisition & Maint.',
    'GRO': 'Food & Food Service',
    'LND': 'Facility Landscaping',
    'MED': 'Health & Medical',
    'MRO': 'Maintenance, Repair & Ops',
    'OFF': 'Office Supplies',
    'PRF': 'Professional Services',
    'PSE': 'Public Safety & Security',
    'SFC': 'Sustainable Facilities',
    'TRD': 'Tradespersons',
    'WMR': 'Waste Mgmt & Recycling',
}

# Sheets to skip (non-data)
SKIP_SHEETS = ['Abbreviations ']

# Color palette — modern and cohesive
PALETTE_MAIN = '#1B2838'       # Dark navy background feel
PALETTE_ACCENT = '#00BCD4'     # Cyan accent
PALETTE_ACCENT2 = '#FF6F61'    # Coral accent
PALETTE_ACCENT3 = '#7C4DFF'    # Purple accent
PALETTE_CAT = sns.color_palette("husl", 15)
PALETTE_IT = ['#00BCD4', '#7C4DFF', '#FF6F61']

# Plot styling
plt.rcParams.update({
    'figure.facecolor': '#FAFAFA',
    'axes.facecolor': '#FAFAFA',
    'axes.edgecolor': '#CCCCCC',
    'axes.labelcolor': '#333333',
    'xtick.color': '#555555',
    'ytick.color': '#555555',
    'font.family': 'sans-serif',
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'figure.titlesize': 16,
    'figure.titleweight': 'bold',
})


# ──────────────────────────────────────────────────────────────────
# Data Loading & Cleaning
# ──────────────────────────────────────────────────────────────────
def load_vendor_data(filepath: str) -> pd.DataFrame:
    """Load and clean vendor contact details from all category sheets."""
    xl = pd.ExcelFile(filepath)
    frames = []

    for sheet in xl.sheet_names:
        if sheet.strip() in [s.strip() for s in SKIP_SHEETS]:
            continue

        df = pd.read_excel(filepath, sheet_name=sheet, header=0)
        # Keep only meaningful columns (first 7)
        df = df.iloc[:, :7]
        df.columns = ['Contract_Code', 'Name', 'Company', 'Role', 'Email', 'Phone', 'SDO_Commitment_Pct']

        # Add category identifier
        df['Category'] = sheet.strip()

        # Remove rows that are headers, metadata, or master contract records
        metadata_keywords = [
            'Master Contract', 'Solicitation Enabled', 'Master MBPO',
            'Bid and Contract', 'Category and Vendor', 'Category Development',
            'Mass Gov', 'OSD Help Desk', 'N/A'
        ]
        mask = df['Company'].astype(str).apply(
            lambda x: not any(kw.lower() in x.lower() for kw in metadata_keywords)
        )
        df = df[mask].copy()

        # Clean SDO column
        df['SDO_Commitment_Pct'] = pd.to_numeric(df['SDO_Commitment_Pct'], errors='coerce')

        # Clean company names
        df['Company'] = df['Company'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
        df['Name'] = df['Name'].astype(str).str.replace('\n', ' ', regex=False).str.strip()

        # Clean contract codes
        df['Contract_Code'] = df['Contract_Code'].astype(str).str.replace('\n', ' ', regex=False).str.strip()

        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    print(f"✅ Loaded vendor data: {len(combined):,} vendor records across {combined['Category'].nunique()} categories")
    return combined


def load_categorized_companies(filepath: str) -> pd.DataFrame:
    """Load and clean categorized companies data."""
    xl = pd.ExcelFile(filepath)
    frames = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
        # Row 0 is sub-header (National & Local, Local Companies, SGC Target)
        # Row 1 is the actual data header with industry name
        # Data starts from row 2

        industry = sheet.strip()

        # Extract the three lists
        national_local = df.iloc[2:, 1].dropna().astype(str).str.strip().tolist()
        local_only = df.iloc[2:, 2].dropna().astype(str).str.strip().tolist()
        sgc_target = df.iloc[2:, 3].dropna().astype(str).str.strip().tolist()

        for company in national_local:
            if company and company != 'nan':
                frames.append({'Industry': industry, 'Company': company, 'Type': 'National & Local'})
        for company in local_only:
            if company and company != 'nan':
                frames.append({'Industry': industry, 'Company': company, 'Type': 'Local'})
        for company in sgc_target:
            if company and company != 'nan':
                frames.append({'Industry': industry, 'Company': company, 'Type': 'SGC Target'})

    combined = pd.DataFrame(frames)
    print(f"✅ Loaded categorized companies: {len(combined):,} entries across {combined['Industry'].nunique()} industries")
    return combined


# ──────────────────────────────────────────────────────────────────
# Helper Functions
# ──────────────────────────────────────────────────────────────────
def save_plot(fig, filename: str):
    """Save figure with high quality."""
    fig.savefig(OUTPUT_DIR / filename, dpi=200, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"   📊 Saved: {filename}")


def get_category_label(code: str) -> str:
    """Get human-readable label for category code."""
    return CATEGORY_NAMES.get(code, code)


def wrap_labels(labels, width=20):
    """Wrap long labels for better readability."""
    return [textwrap.fill(str(l), width) for l in labels]


# ──────────────────────────────────────────────────────────────────
# BQ1: Top IT Sector Companies by SDO Commitment %
# ──────────────────────────────────────────────────────────────────
def bq1_it_sector_top_sdo(vendor_df: pd.DataFrame):
    """Which companies in the IT sector have the best SDO commitment percentage?"""
    print("\n" + "="*70)
    print("BQ1: Top IT Sector Companies by SDO Commitment %")
    print("="*70)

    it_df = vendor_df[vendor_df['Category'].isin(['ITE', 'ITS', 'ITT'])].copy()
    it_with_sdo = it_df.dropna(subset=['SDO_Commitment_Pct'])
    it_with_sdo = it_with_sdo[it_with_sdo['SDO_Commitment_Pct'] > 0]

    # Aggregate by company (take max SDO across contracts)
    company_sdo = it_with_sdo.groupby(['Company', 'Category']).agg(
        SDO_Max=('SDO_Commitment_Pct', 'max'),
        Contract_Count=('Contract_Code', 'nunique')
    ).reset_index().sort_values('SDO_Max', ascending=False)

    top15 = company_sdo.head(15)

    # Statistics
    print(f"\n  Total IT vendors with SDO data: {len(it_with_sdo):,}")
    print(f"  Unique companies with SDO: {company_sdo['Company'].nunique()}")
    print(f"  SDO Range: {company_sdo['SDO_Max'].min():.2%} – {company_sdo['SDO_Max'].max():.2%}")
    print(f"  Mean SDO: {company_sdo['SDO_Max'].mean():.2%}")
    print(f"  Median SDO: {company_sdo['SDO_Max'].median():.2%}")
    print(f"\n  Top 5 Companies:")
    for _, row in top15.head().iterrows():
        print(f"    {row['Company']:<45s} {row['SDO_Max']:.0%}  ({row['Category']})")

    # ── Plot ──
    fig, ax = plt.subplots(figsize=(12, 8))

    colors = [PALETTE_IT[['ITE', 'ITS', 'ITT'].index(cat)] for cat in top15['Category']]
    bars = ax.barh(range(len(top15)), top15['SDO_Max'].values * 100, color=colors, edgecolor='white', linewidth=0.5)

    ax.set_yticks(range(len(top15)))
    ax.set_yticklabels([textwrap.fill(c, 35) for c in top15['Company']], fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel('SDO Commitment (%)', fontweight='bold')
    ax.set_title('BQ1: Top 15 IT Sector Companies by SDO Commitment %', pad=15)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter())

    # Add value labels
    for bar, val in zip(bars, top15['SDO_Max'].values):
        ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
                f'{val:.0%}', va='center', fontsize=9, fontweight='bold', color='#333')

    # Legend
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor=PALETTE_IT[0], label='ITE – IT Equipment'),
        Patch(facecolor=PALETTE_IT[1], label='ITS – IT Software'),
        Patch(facecolor=PALETTE_IT[2], label='ITT – Telecom'),
    ]
    ax.legend(handles=legend_elements, loc='lower right', framealpha=0.9)

    fig.tight_layout()
    save_plot(fig, 'BQ1_IT_Top_SDO_Companies.png')


# ──────────────────────────────────────────────────────────────────
# BQ2: Average SDO Commitment % by Category
# ──────────────────────────────────────────────────────────────────
def bq2_avg_sdo_by_category(vendor_df: pd.DataFrame):
    """Which industry/category has the highest average SDO commitment percentage?"""
    print("\n" + "="*70)
    print("BQ2: Average SDO Commitment % by Procurement Category")
    print("="*70)

    sdo_df = vendor_df.dropna(subset=['SDO_Commitment_Pct'])
    sdo_df = sdo_df[sdo_df['SDO_Commitment_Pct'] > 0]

    # Cap extreme outliers at 1.0 (100%) for meaningful comparison
    sdo_capped = sdo_df.copy()
    sdo_capped['SDO_Commitment_Pct'] = sdo_capped['SDO_Commitment_Pct'].clip(upper=1.0)

    cat_stats = sdo_capped.groupby('Category').agg(
        Avg_SDO=('SDO_Commitment_Pct', 'mean'),
        Median_SDO=('SDO_Commitment_Pct', 'median'),
        Count=('SDO_Commitment_Pct', 'count'),
        Std=('SDO_Commitment_Pct', 'std')
    ).reset_index().sort_values('Avg_SDO', ascending=True)

    cat_stats['Label'] = cat_stats['Category'].map(get_category_label)

    print(f"\n  {'Category':<35s} {'Avg SDO':>10s} {'Median':>10s} {'Count':>8s}")
    print("  " + "-"*65)
    for _, row in cat_stats.sort_values('Avg_SDO', ascending=False).iterrows():
        print(f"  {row['Label']:<35s} {row['Avg_SDO']:>9.2%} {row['Median_SDO']:>9.2%} {row['Count']:>8.0f}")

    # ── Plot ──
    fig, ax = plt.subplots(figsize=(12, 8))

    colors = sns.color_palette("viridis", len(cat_stats))
    bars = ax.barh(range(len(cat_stats)), cat_stats['Avg_SDO'].values * 100, color=colors, edgecolor='white')

    ax.set_yticks(range(len(cat_stats)))
    ax.set_yticklabels(cat_stats['Label'].values, fontsize=9)
    ax.set_xlabel('Average SDO Commitment (%)', fontweight='bold')
    ax.set_title('BQ2: Average SDO Commitment % by Procurement Category', pad=15)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter())

    for bar, val, count in zip(bars, cat_stats['Avg_SDO'].values, cat_stats['Count'].values):
        ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
                f'{val:.1%} (n={count:.0f})', va='center', fontsize=8.5, color='#333')

    fig.tight_layout()
    save_plot(fig, 'BQ2_Avg_SDO_by_Category.png')


# ──────────────────────────────────────────────────────────────────
# BQ3: Distribution of Vendors Across Procurement Categories
# ──────────────────────────────────────────────────────────────────
def bq3_vendor_distribution(vendor_df: pd.DataFrame):
    """What is the distribution of vendors across procurement categories?"""
    print("\n" + "="*70)
    print("BQ3: Distribution of Vendors Across Procurement Categories")
    print("="*70)

    cat_counts = vendor_df.groupby('Category')['Company'].nunique().sort_values(ascending=False)
    cat_counts.index = cat_counts.index.map(get_category_label)

    total = cat_counts.sum()
    print(f"\n  Total unique vendors: {total:,}")
    for cat, count in cat_counts.items():
        print(f"  {cat:<35s} {count:>5d} vendors  ({count/total:>6.1%})")

    # ── Plot: Treemap using matplotlib (squarify-style) ──
    fig, ax = plt.subplots(figsize=(14, 9))

    sizes = cat_counts.values
    labels = [f"{cat}\n{count:,} vendors\n({count/total:.1%})" for cat, count in cat_counts.items()]
    colors = sns.color_palette("Set3", len(sizes))

    # Simple treemap via nested rectangles (squarify algorithm)
    try:
        import squarify
        squarify.plot(sizes=sizes, label=labels, color=colors, alpha=0.85,
                      text_kwargs={'fontsize': 9, 'fontweight': 'bold'}, ax=ax)
        ax.axis('off')
    except ImportError:
        # Fallback to horizontal bar
        ax.barh(range(len(cat_counts)), cat_counts.values, color=colors, edgecolor='white')
        ax.set_yticks(range(len(cat_counts)))
        ax.set_yticklabels(cat_counts.index, fontsize=9)
        ax.set_xlabel('Number of Unique Vendors', fontweight='bold')

    ax.set_title('BQ3: Distribution of Vendors Across Procurement Categories', pad=15, fontsize=14, fontweight='bold')

    fig.tight_layout()
    save_plot(fig, 'BQ3_Vendor_Distribution_Treemap.png')


# ──────────────────────────────────────────────────────────────────
# BQ4: Vendor Count by Contract Sub-Category Codes
# ──────────────────────────────────────────────────────────────────
def bq4_contract_subcategories(vendor_df: pd.DataFrame):
    """How are vendors distributed across contract sub-category codes?"""
    print("\n" + "="*70)
    print("BQ4: Vendor Count by Contract Sub-Category Codes")
    print("="*70)

    # Extract the base contract code (e.g., ITS55, FAC112)
    code_counts = vendor_df.groupby('Contract_Code')['Company'].nunique()
    code_counts = code_counts[code_counts.index.str.len() <= 15]  # filter out garbage
    code_counts = code_counts.sort_values(ascending=False).head(25)

    print(f"\n  Top 25 contract codes by vendor count:")
    for code, count in code_counts.head(15).items():
        print(f"    {code:<20s} {count:>5d} vendors")

    # ── Plot ──
    fig, ax = plt.subplots(figsize=(14, 8))

    colors = sns.color_palette("coolwarm", len(code_counts))
    bars = ax.bar(range(len(code_counts)), code_counts.values, color=colors, edgecolor='white', linewidth=0.5)

    ax.set_xticks(range(len(code_counts)))
    ax.set_xticklabels(code_counts.index, rotation=45, ha='right', fontsize=8)
    ax.set_ylabel('Number of Unique Vendors', fontweight='bold')
    ax.set_xlabel('Contract Sub-Category Code', fontweight='bold')
    ax.set_title('BQ4: Top 25 Contract Sub-Categories by Vendor Count', pad=15)

    # Add value labels on top
    for bar, val in zip(bars, code_counts.values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                str(val), ha='center', va='bottom', fontsize=8, fontweight='bold')

    fig.tight_layout()
    save_plot(fig, 'BQ4_Contract_Subcategory_Vendors.png')


# ──────────────────────────────────────────────────────────────────
# BQ5: National vs Local Company Presence Across Industries
# ──────────────────────────────────────────────────────────────────
def bq5_national_vs_local(cat_df: pd.DataFrame):
    """How do National vs Local companies compare across industries?"""
    print("\n" + "="*70)
    print("BQ5: National vs Local Company Presence Across Industries")
    print("="*70)

    pivot = cat_df.groupby(['Industry', 'Type']).size().unstack(fill_value=0)
    pivot = pivot.sort_values(by='National & Local', ascending=True)

    print(f"\n  {'Industry':<35s} {'National&Local':>15s} {'Local':>10s} {'SGC Target':>12s}")
    print("  " + "-"*75)
    for idx, row in pivot.iterrows():
        print(f"  {idx:<35s} {row.get('National & Local', 0):>15.0f} {row.get('Local', 0):>10.0f} {row.get('SGC Target', 0):>12.0f}")

    # ── Plot: Stacked horizontal bar ──
    fig, ax = plt.subplots(figsize=(13, 8))

    types = ['National & Local', 'Local', 'SGC Target']
    colors_stack = ['#00BCD4', '#7C4DFF', '#FF6F61']
    bottom = np.zeros(len(pivot))

    for typ, color in zip(types, colors_stack):
        if typ in pivot.columns:
            values = pivot[typ].values
            ax.barh(range(len(pivot)), values, left=bottom, color=color, label=typ, edgecolor='white', linewidth=0.5)
            bottom += values

    ax.set_yticks(range(len(pivot)))
    ax.set_yticklabels([textwrap.fill(str(i), 25) for i in pivot.index], fontsize=9)
    ax.set_xlabel('Number of Companies', fontweight='bold')
    ax.set_title('BQ5: National vs Local Company Presence Across Industries', pad=15)
    ax.legend(loc='lower right', framealpha=0.9)

    fig.tight_layout()
    save_plot(fig, 'BQ5_National_vs_Local_Industries.png')


# ──────────────────────────────────────────────────────────────────
# BQ6: SDO Commitment Distribution & Outlier Detection
# ──────────────────────────────────────────────────────────────────
def bq6_sdo_distribution_outliers(vendor_df: pd.DataFrame):
    """What is the SDO commitment distribution? Are there outliers?"""
    print("\n" + "="*70)
    print("BQ6: SDO Commitment Distribution & Outlier Detection")
    print("="*70)

    sdo_df = vendor_df.dropna(subset=['SDO_Commitment_Pct']).copy()
    sdo_df = sdo_df[sdo_df['SDO_Commitment_Pct'] > 0]

    # Cap at 1.0 for visualization (values above are likely data issues)
    sdo_df['SDO_Capped'] = sdo_df['SDO_Commitment_Pct'].clip(upper=1.0)

    # Categories with enough data
    cat_counts = sdo_df.groupby('Category').size()
    valid_cats = cat_counts[cat_counts >= 10].index.tolist()
    sdo_valid = sdo_df[sdo_df['Category'].isin(valid_cats)]

    # Outlier detection using IQR
    for cat in valid_cats:
        cat_data = sdo_valid[sdo_valid['Category'] == cat]['SDO_Capped']
        q1, q3 = cat_data.quantile(0.25), cat_data.quantile(0.75)
        iqr = q3 - q1
        outliers = cat_data[(cat_data < q1 - 1.5*iqr) | (cat_data > q3 + 1.5*iqr)]
        if len(outliers) > 0:
            print(f"  {get_category_label(cat)}: {len(outliers)} outliers detected (IQR method)")

    # ── Plot: Box plot ──
    fig, ax = plt.subplots(figsize=(14, 8))

    plot_data = sdo_valid[['Category', 'SDO_Capped']].copy()
    plot_data['Label'] = plot_data['Category'].map(get_category_label)

    # Order by median
    order = plot_data.groupby('Label')['SDO_Capped'].median().sort_values(ascending=False).index.tolist()

    box = sns.boxplot(data=plot_data, x='Label', y='SDO_Capped', order=order,
                      palette="husl", ax=ax, fliersize=4, linewidth=1.2)

    ax.set_xticklabels(wrap_labels(order, 15), rotation=45, ha='right', fontsize=8)
    ax.set_ylabel('SDO Commitment (%) — Capped at 100%', fontweight='bold')
    ax.set_xlabel('')
    ax.set_title('BQ6: SDO Commitment Distribution by Category (with Outliers)', pad=15)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=1.0))

    fig.tight_layout()
    save_plot(fig, 'BQ6_SDO_Distribution_Boxplot.png')


# ──────────────────────────────────────────────────────────────────
# BQ7: SDO Coverage Rate by Category
# ──────────────────────────────────────────────────────────────────
def bq7_sdo_coverage_rate(vendor_df: pd.DataFrame):
    """What proportion of vendors have valid SDO commitments?"""
    print("\n" + "="*70)
    print("BQ7: SDO Coverage Rate — Valid vs Missing SDO")
    print("="*70)

    coverage = vendor_df.groupby('Category').apply(
        lambda g: pd.Series({
            'Has_SDO': g['SDO_Commitment_Pct'].notna().sum(),
            'No_SDO': g['SDO_Commitment_Pct'].isna().sum(),
            'Total': len(g),
            'Coverage_Rate': g['SDO_Commitment_Pct'].notna().mean()
        })
    ).reset_index()
    coverage['Label'] = coverage['Category'].map(get_category_label)
    coverage = coverage.sort_values('Coverage_Rate', ascending=True)

    print(f"\n  {'Category':<35s} {'Has SDO':>10s} {'No SDO':>10s} {'Coverage':>10s}")
    print("  " + "-"*68)
    for _, row in coverage.sort_values('Coverage_Rate', ascending=False).iterrows():
        print(f"  {row['Label']:<35s} {row['Has_SDO']:>10.0f} {row['No_SDO']:>10.0f} {row['Coverage_Rate']:>9.1%}")

    # ── Plot: Grouped bar ──
    fig, ax = plt.subplots(figsize=(14, 8))

    x = np.arange(len(coverage))
    width = 0.35

    ax.barh(x - width/2, coverage['Has_SDO'].values, width, color='#00BCD4', label='Has SDO Commitment', edgecolor='white')
    ax.barh(x + width/2, coverage['No_SDO'].values, width, color='#FF6F61', label='No SDO (Missing/N/A)', edgecolor='white')

    ax.set_yticks(x)
    ax.set_yticklabels(coverage['Label'].values, fontsize=9)
    ax.set_xlabel('Number of Vendors', fontweight='bold')
    ax.set_title('BQ7: SDO Coverage Rate — Vendors with Valid SDO vs Missing', pad=15)
    ax.legend(loc='lower right', framealpha=0.9)

    # Add coverage % annotation
    for i, (_, row) in enumerate(coverage.iterrows()):
        ax.text(max(row['Has_SDO'], row['No_SDO']) + 5, i,
                f"{row['Coverage_Rate']:.0%}", va='center', fontsize=9, fontweight='bold', color='#333')

    fig.tight_layout()
    save_plot(fig, 'BQ7_SDO_Coverage_Rate.png')


# ──────────────────────────────────────────────────────────────────
# BQ8: Correlation — Vendor Count vs Avg SDO
# ──────────────────────────────────────────────────────────────────
def bq8_vendor_count_vs_sdo(vendor_df: pd.DataFrame):
    """Is there a correlation between the number of vendors and average SDO?"""
    print("\n" + "="*70)
    print("BQ8: Correlation — Vendor Count vs Avg SDO by Category")
    print("="*70)

    sdo_df = vendor_df.dropna(subset=['SDO_Commitment_Pct'])
    sdo_df = sdo_df[sdo_df['SDO_Commitment_Pct'] > 0]
    sdo_df = sdo_df[sdo_df['SDO_Commitment_Pct'] <= 1.0]  # Filter unreasonable values

    cat_stats = sdo_df.groupby('Category').agg(
        Vendor_Count=('Company', 'nunique'),
        Avg_SDO=('SDO_Commitment_Pct', 'mean')
    ).reset_index()
    cat_stats['Label'] = cat_stats['Category'].map(get_category_label)

    # Correlation
    from scipy import stats
    if len(cat_stats) > 2:
        corr, pval = stats.pearsonr(cat_stats['Vendor_Count'], cat_stats['Avg_SDO'])
        print(f"\n  Pearson correlation: r = {corr:.3f}, p-value = {pval:.4f}")
        print(f"  Interpretation: {'Significant' if pval < 0.05 else 'Not significant'} at α=0.05")

    # ── Plot: Scatter with regression ──
    fig, ax = plt.subplots(figsize=(10, 8))

    ax.scatter(cat_stats['Vendor_Count'], cat_stats['Avg_SDO'] * 100,
               s=150, c=PALETTE_CAT[:len(cat_stats)], edgecolors='white', linewidth=1.5, zorder=5)

    # Add labels
    for _, row in cat_stats.iterrows():
        ax.annotate(row['Category'],
                    (row['Vendor_Count'], row['Avg_SDO'] * 100),
                    textcoords="offset points", xytext=(8, 5), fontsize=9, fontweight='bold')

    # Regression line
    if len(cat_stats) > 2:
        z = np.polyfit(cat_stats['Vendor_Count'], cat_stats['Avg_SDO'] * 100, 1)
        p = np.poly1d(z)
        x_line = np.linspace(cat_stats['Vendor_Count'].min(), cat_stats['Vendor_Count'].max(), 100)
        ax.plot(x_line, p(x_line), '--', color='#FF6F61', linewidth=2, alpha=0.7, label=f'r = {corr:.3f}')
        ax.legend(fontsize=11)

    ax.set_xlabel('Number of Unique Vendors', fontweight='bold')
    ax.set_ylabel('Average SDO Commitment (%)', fontweight='bold')
    ax.set_title('BQ8: Vendor Count vs Avg SDO Commitment by Category', pad=15)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())

    fig.tight_layout()
    save_plot(fig, 'BQ8_Vendor_Count_vs_SDO_Scatter.png')


# ──────────────────────────────────────────────────────────────────
# BQ9: Industry Diversity Heatmap (Categorized Companies)
# ──────────────────────────────────────────────────────────────────
def bq9_industry_diversity_heatmap(cat_df: pd.DataFrame):
    """How diverse is the industry representation? What's the SGC target gap?"""
    print("\n" + "="*70)
    print("BQ9: Industry Diversity — National vs Local vs SGC Target")
    print("="*70)

    pivot = cat_df.groupby(['Industry', 'Type']).size().unstack(fill_value=0)

    print(f"\n  Company counts per industry and type:")
    print(pivot.to_string())

    # ── Plot: Heatmap ──
    fig, ax = plt.subplots(figsize=(10, 8))

    # Wrap long industry names
    pivot.index = [textwrap.fill(str(i), 22) for i in pivot.index]

    sns.heatmap(pivot, annot=True, fmt='d', cmap='YlOrRd', linewidths=1, linecolor='white',
                ax=ax, cbar_kws={'label': 'Number of Companies'})

    ax.set_title('BQ9: Industry Company Counts — National vs Local vs SGC Target', pad=15)
    ax.set_ylabel('')
    ax.set_xlabel('')

    fig.tight_layout()
    save_plot(fig, 'BQ9_Industry_Diversity_Heatmap.png')


# ──────────────────────────────────────────────────────────────────
# BQ10: Vendor Concentration in IT Sector
# ──────────────────────────────────────────────────────────────────
def bq10_it_vendor_concentration(vendor_df: pd.DataFrame):
    """Do a few companies dominate the IT sector?"""
    print("\n" + "="*70)
    print("BQ10: Vendor Concentration in IT Sector")
    print("="*70)

    it_df = vendor_df[vendor_df['Category'].isin(['ITE', 'ITS', 'ITT'])].copy()

    # Count contracts per company
    company_contracts = it_df.groupby('Company')['Contract_Code'].nunique().sort_values(ascending=False)

    total_contracts = company_contracts.sum()
    top10 = company_contracts.head(10)
    rest = total_contracts - top10.sum()

    print(f"\n  Total IT companies: {len(company_contracts):,}")
    print(f"  Total contract associations: {total_contracts:,}")
    print(f"  Top 10 companies hold: {top10.sum()/total_contracts:.1%} of contracts")
    print(f"\n  Top 10 Companies:")
    for company, count in top10.items():
        print(f"    {company:<45s} {count:>3d} contracts ({count/total_contracts:.1%})")

    # ── Plot: Donut chart ──
    fig, ax = plt.subplots(figsize=(10, 10))

    sizes = list(top10.values) + [rest]
    labels = [textwrap.fill(c, 25) for c in top10.index] + [f'Other ({len(company_contracts)-10} companies)']
    colors = sns.color_palette("husl", len(sizes))

    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct='%1.1f%%', startangle=90,
        colors=colors, pctdistance=0.82, wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2)
    )

    for autotext in autotexts:
        autotext.set_fontsize(8)
        autotext.set_fontweight('bold')

    ax.legend(labels, loc='center left', bbox_to_anchor=(1, 0.5), fontsize=8, framealpha=0.9)
    ax.set_title('BQ10: IT Sector Vendor Concentration\n(Top 10 Companies vs Rest)', pad=20, fontsize=14, fontweight='bold')

    # Center text
    ax.text(0, 0, f'{len(company_contracts)}\nCompanies', ha='center', va='center',
            fontsize=16, fontweight='bold', color='#333')

    fig.tight_layout()
    save_plot(fig, 'BQ10_IT_Vendor_Concentration_Donut.png')


# ──────────────────────────────────────────────────────────────────
# Main Execution
# ──────────────────────────────────────────────────────────────────
def main():
    print("╔══════════════════════════════════════════════════════════════════════╗")
    print("║   ALY 6980 Capstone — Exploratory Data Analysis                    ║")
    print("║   Massachusetts Open Checkbook: Vendor & SDO Analysis              ║")
    print("╚══════════════════════════════════════════════════════════════════════╝")

    # Load data
    vendor_df = load_vendor_data(VENDOR_FILE)
    cat_df = load_categorized_companies(CATEGORIZED_FILE)

    # Run all business question analyses
    bq1_it_sector_top_sdo(vendor_df)
    bq2_avg_sdo_by_category(vendor_df)
    bq3_vendor_distribution(vendor_df)
    bq4_contract_subcategories(vendor_df)
    bq5_national_vs_local(cat_df)
    bq6_sdo_distribution_outliers(vendor_df)
    bq7_sdo_coverage_rate(vendor_df)
    bq8_vendor_count_vs_sdo(vendor_df)
    bq9_industry_diversity_heatmap(cat_df)
    bq10_it_vendor_concentration(vendor_df)

    print("\n" + "="*70)
    print(f"✅ EDA Complete! All plots saved to: {OUTPUT_DIR.resolve()}")
    print("="*70)


if __name__ == '__main__':
    main()
