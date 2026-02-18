import dashboard
import plotly.io as pio
import textwrap

# Set default template to match dashboard
pio.templates.default = "plotly_dark"

print("🚀 Starting Static Dashboard Build...")

# ──────────────────────────────────────────────────────────────────
# HTML Template
# ──────────────────────────────────────────────────────────────────
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MA Open Checkbook Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: #0F172A;
            color: #F1F5F9;
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            padding-bottom: 40px;
        }
        
        /* Layout */
        .container { max-width: 1400px; margin: 0 auto; }
        
        /* Header */
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px 32px;
            border-bottom: 1px solid #334155;
            background: #0F172A;
        }
        h1 {
            font-size: 28px;
            font-weight: 700;
            background: linear-gradient(135deg, #06B6D4, #8B5CF6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .subtitle { color: #94A3B8; font-size: 14px; margin-top: 4px; }
        .badge {
            background: #1E293B; color: #06B6D4;
            padding: 8px 16px; border-radius: 20px; font-size: 12px;
            border: 1px solid #334155; font-weight: 600;
        }

        /* KPI Cards */
        .kpi-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            padding: 24px 32px;
        }
        .kpi-card {
            background: #1E293B;
            border-radius: 12px;
            padding: 24px;
            border: 1px solid #334155;
            text-align: center;
        }
        .kpi-title { color: #94A3B8; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; font-weight: 500; margin-bottom: 8px; }
        .kpi-value { font-size: 32px; font-weight: 700; margin-bottom: 4px; }
        .kpi-sub { color: #94A3B8; font-size: 12px; }

        /* Tabs */
        .tabs {
            display: flex;
            padding: 0 32px;
            border-bottom: 1px solid #334155;
            margin-bottom: 24px;
        }
        .tab-btn {
            background: transparent;
            color: #94A3B8;
            border: none;
            padding: 16px 24px;
            font-size: 14px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            cursor: pointer;
            transition: all 0.2s;
            border-bottom: 3px solid transparent;
        }
        .tab-btn:hover { color: #F1F5F9; }
        .tab-btn.active {
            color: #06B6D4;
            border-bottom: 3px solid #06B6D4;
        }

        /* Tab Content */
        .tab-content {
            display: none;
            padding: 0 32px;
            animation: fadeIn 0.3s ease-in-out;
        }
        .tab-content.active { display: block; }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        /* Charts */
        .chart-row {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
        }
        .chart-col { flex: 1; min-width: 0; } /* min-width 0 fixes flex overflow */
        .chart-card {
            background: #1E293B;
            border-radius: 12px;
            padding: 16px;
            border: 1px solid #334155;
            height: 100%;
        }

        /* Controls */
        .controls {
            display: flex;
            align-items: center;
            margin-bottom: 20px;
            margin-top: 10px;
        }
        label { color: #94A3B8; font-size: 13px; font-weight: 500; margin-right: 12px; }
        select {
            background: #1E293B;
            color: #F1F5F9;
            border: 1px solid #334155;
            padding: 8px 12px;
            border-radius: 8px;
            font-size: 14px;
            width: 300px;
            outline: none;
        }

        /* Footer */
        footer {
            margin-top: 40px;
            padding: 24px;
            text-align: center;
            border-top: 1px solid #334155;
            color: #94A3B8;
            font-size: 12px;
        }
    </style>
</head>
<body>

<div class="container">
    <header>
        <div>
            <h1>Massachusetts Open Checkbook</h1>
            <div class="subtitle">Vendor Contract & SDO Commitment Analysis Dashboard</div>
        </div>
        <div>
            <span class="badge">ALY 6980 Capstone</span>
        </div>
    </header>

    <!-- KPI Row -->
    <div class="kpi-row">
        {kpi_html}
    </div>

    <!-- Navigation -->
    <div class="tabs">
        <button class="tab-btn active" onclick="openTab(event, 'tab-1')">🖥️ IT Sector SDO</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-2')">📊 Cross-Category</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-3')">✅ Vendor Coverage</button>
        <button class="tab-btn" onclick="openTab(event, 'tab-4')">🏢 Industry Analysis</button>
    </div>

    <!-- Tab 1: IT Sector -->
    <div id="tab-1" class="tab-content active">
        <div class="controls">
            <label>Filter by IT Sub-Category:</label>
            <select id="it-filter" onchange="updateItChart()">
                <option value="All">All IT Categories</option>
                <option value="ITE">ITE — IT Equipment & Services</option>
                <option value="ITS">ITS — IT Software & Services</option>
                <option value="ITT">ITT — Telecom & Networking</option>
            </select>
        </div>
        
        <div class="chart-row">
            <div class="chart-col">
                <div class="chart-card">
                    <!-- Pre-generated charts for BQ1 variants -->
                    <div id="bq1-All" class="bq1-variant">{bq1_all}</div>
                    <div id="bq1-ITE" class="bq1-variant" style="display:none;">{bq1_ite}</div>
                    <div id="bq1-ITS" class="bq1-variant" style="display:none;">{bq1_its}</div>
                    <div id="bq1-ITT" class="bq1-variant" style="display:none;">{bq1_itt}</div>
                </div>
            </div>
        </div>
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq10}</div></div>
            <div class="chart-col"><div class="chart-card">{it_radar}</div></div>
        </div>
    </div>

    <!-- Tab 2: Cross-Category -->
    <div id="tab-2" class="tab-content">
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq2}</div></div>
        </div>
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq6}</div></div>
            <div class="chart-col"><div class="chart-card">{bq8}</div></div>
        </div>
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{sdo_hist}</div></div>
        </div>
    </div>

    <!-- Tab 3: Vendor Coverage -->
    <div id="tab-3" class="tab-content">
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq7}</div></div>
        </div>
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq3}</div></div>
            <div class="chart-col"><div class="chart-card">{bq4}</div></div>
        </div>
    </div>

    <!-- Tab 4: Industry Analysis -->
    <div id="tab-4" class="tab-content">
        <div class="chart-row">
            <div class="chart-col"><div class="chart-card">{bq5}</div></div>
            <div class="chart-col"><div class="chart-card">{bq9}</div></div>
        </div>
    </div>

    <footer>
        <p>ALY 6980 Capstone Project • Sumesh Chakkaravarthi • February 2026</p>
    </footer>
</div>

<script>
    // Tab Switching Logic
    function openTab(evt, tabName) {
        var i, tabcontent, tablinks;
        tabcontent = document.getElementsByClassName("tab-content");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].classList.remove("active");
        }
        tablinks = document.getElementsByClassName("tab-btn");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].classList.remove("active");
        }
        document.getElementById(tabName).classList.add("active");
        evt.currentTarget.classList.add("active");
        
        // Trigger plot resize to fix layout bugs when unhiding
        window.dispatchEvent(new Event('resize'));
    }

    // Dropdown Filter Logic for BQ1
    function updateItChart() {
        var val = document.getElementById("it-filter").value;
        var variants = document.getElementsByClassName("bq1-variant");
        for (var i = 0; i < variants.length; i++) {
            variants[i].style.display = "none";
        }
        document.getElementById("bq1-" + val).style.display = "block";
    }
</script>

</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────
# Build Process
# ──────────────────────────────────────────────────────────────────

def generate_kpi_html():
    kpis = [
        ("Total Vendors", f"{dashboard.total_vendors:,}", "Unique companies", dashboard.COLORS['accent']),
        ("Average SDO", f"{dashboard.avg_sdo:.1%}", "Commitment percentage", dashboard.COLORS['accent2']),
        ("Categories", str(dashboard.categories_count), "Procurement categories", dashboard.COLORS['accent3']),
        ("IT Vendors", f"{dashboard.it_vendors:,}", "ITE + ITS + ITT", dashboard.COLORS['accent4']),
        ("Industries", str(dashboard.industries_count), "Company classifications", '#F59E0B'),
    ]
    html_parts = []
    for title, val, sub, color in kpis:
        card = f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value" style="color: {color}">{val}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """
        html_parts.append(card)
    return "\n".join(html_parts)

def fig_to_html(fig):
    return pio.to_html(fig, include_plotlyjs=False, full_html=False, config={'displayModeBar': False})

print("Generating Charts...")

# KPI
kpi_html_content = generate_kpi_html()

# Tab 1: IT Sector
print(" - Building Tab 1 (including filter variants)...")
bq1_all = fig_to_html(dashboard.make_bq1_chart('All'))
bq1_ite = fig_to_html(dashboard.make_bq1_chart('ITE'))
bq1_its = fig_to_html(dashboard.make_bq1_chart('ITS'))
bq1_itt = fig_to_html(dashboard.make_bq1_chart('ITT'))
bq10 = fig_to_html(dashboard.make_bq10_chart())
it_radar = fig_to_html(dashboard.make_it_radar())

# Tab 2: Cross-Category
print(" - Building Tab 2...")
bq2 = fig_to_html(dashboard.make_bq2_chart())
bq6 = fig_to_html(dashboard.make_bq6_chart())
bq8 = fig_to_html(dashboard.make_bq8_chart())
sdo_hist = fig_to_html(dashboard.make_sdo_histogram())

# Tab 3: Coverage
print(" - Building Tab 3...")
bq7 = fig_to_html(dashboard.make_bq7_chart())
bq3 = fig_to_html(dashboard.make_bq3_chart())
bq4 = fig_to_html(dashboard.make_bq4_chart())

# Tab 4: Industry
print(" - Building Tab 4...")
bq5 = fig_to_html(dashboard.make_bq5_chart())
bq9 = fig_to_html(dashboard.make_bq9_chart())

print("Assembling HTML...")
final_html = HTML_TEMPLATE
final_html = final_html.replace("{kpi_html}", kpi_html_content)
final_html = final_html.replace("{bq1_all}", bq1_all)
final_html = final_html.replace("{bq1_ite}", bq1_ite)
final_html = final_html.replace("{bq1_its}", bq1_its)
final_html = final_html.replace("{bq1_itt}", bq1_itt)
final_html = final_html.replace("{bq10}", bq10)
final_html = final_html.replace("{it_radar}", it_radar)
final_html = final_html.replace("{bq2}", bq2)
final_html = final_html.replace("{bq6}", bq6)
final_html = final_html.replace("{bq8}", bq8)
final_html = final_html.replace("{sdo_hist}", sdo_hist)
final_html = final_html.replace("{bq7}", bq7)
final_html = final_html.replace("{bq3}", bq3)
final_html = final_html.replace("{bq4}", bq4)
final_html = final_html.replace("{bq5}", bq5)
final_html = final_html.replace("{bq9}", bq9)

with open("index.html", "w") as f:
    f.write(final_html)

print("✅ Success! Generated 'index.html'. Open this file in your browser to test.")
