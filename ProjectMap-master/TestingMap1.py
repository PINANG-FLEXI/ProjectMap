import folium
import geopandas as gpd
import pandas as pd
import branca.colormap as cm
from folium.plugins import Fullscreen
import os
from branca.element import Figure, Html, MacroElement
import json

# === 1. Load Shapefile ===
shapefile_path = r"C:/Users/User/Downloads/ProjectMap-master/BATAS PROVINSI DESEMBER 2019 DUKCAPIL/BATAS_PROVINSI_DESEMBER_2019_DUKCAPIL.shp"
gdf = gpd.read_file(shapefile_path)
gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.01, preserve_topology=True)

# === 2. Load Excel ===
data_path = "data_pencairan1.xlsx"
nilai_df = pd.read_excel(data_path)

# Pastikan kolom benar-benar ada
required_cols = [
    "PROVINSI",
    "total_loan_created",
    "total_os_loan_created",
    "total_ktp_reject",
    "total_usr_reject",
    "total_lepas_ktp_reject",   # 🔵 OS KTP
    "total_lepas_usr_reject",   # 🔵 OS USR
    "total_reject_slik",
    "total_reject_sicd_raya",
    "total_reject_sicd_bri",
    "total_reject_blacklist_company",
    "total_reject_score_500",
    "total_reject_failed_to_approve"
]

missing = [c for c in required_cols if c not in nilai_df.columns]
if missing:
    print("⚠️ Missing columns:", missing)

# === 3. Group by PROVINSI (AGGREGATE DATA YANG BENAR) ===
nilai_df_grouped = nilai_df.groupby("PROVINSI").agg({
    "total_loan_created": "sum",
    "total_os_loan_created": "sum",
    "total_ktp_reject": "sum",
    "total_usr_reject": "sum",
    "total_ktp_reject": "sum",          # NOA KTP
    "total_usr_reject": "sum",          # NOA USR
    "total_lepas_ktp_reject": "sum",    # OS KTP
    "total_lepas_usr_reject": "sum",    # OS USR
    "total_reject_slik": "sum",
    "total_reject_sicd_raya": "sum",
    "total_reject_sicd_bri": "sum",
    "total_reject_blacklist_company": "sum",
    "total_reject_score_500": "sum",
    "total_reject_failed_to_approve": "sum"
}).reset_index()

# === 4. Merge hanya sekali (BENAR) ===
gdf = gdf.merge(nilai_df_grouped, how="left", on="PROVINSI")

for col in [
    "total_loan_created", "total_os_loan_created", 
    "total_ktp_reject", "total_usr_reject",
    "total_lepas_ktp_reject", "total_lepas_usr_reject",
    "total_reject_slik", "total_reject_sicd_raya", "total_reject_sicd_bri",
    "total_reject_blacklist_company", "total_reject_score_500",
    "total_reject_failed_to_approve"
]:
    gdf[col] = gdf[col].fillna(0)

# === 5. Hitung Total Reject ===
gdf["total_reject"] = (
    gdf["total_reject_slik"].fillna(0) +
    gdf["total_reject_sicd_raya"].fillna(0) +
    gdf["total_reject_sicd_bri"].fillna(0) +
    gdf["total_reject_blacklist_company"].fillna(0) +
    gdf["total_reject_score_500"].fillna(0) +
    gdf["total_reject_failed_to_approve"].fillna(0)
)

# === 6. Siapkan OS Potensi (KTP & USR) ===
gdf["os_potensi_ktp"] = gdf["total_ktp_reject"].fillna(0)
gdf["os_potensi_usr"] = gdf["total_usr_reject"].fillna(0)
gdf["os_potensi"] = gdf["os_potensi_ktp"] + gdf["os_potensi_usr"]

gdf["os_loan_created"] = gdf["total_os_loan_created"].fillna(0)

# === Format Rupiah ===
def format_rupiah(x):
    try:
        return "Rp {:,.0f}".format(x).replace(",", ".")
    except:
        return "Rp 0"

# === 7. Siapkan Data Chart Top 10 (OS Loan Created) ===
loan_sorted = (
    gdf[["PROVINSI", "total_loan_created", "total_os_loan_created"]]
    .sort_values("total_os_loan_created", ascending=False)
    .head(10)
)

loan_labels = [
    f"{prov} (NOA: {noa})"
    for prov, noa in zip(
        loan_sorted["PROVINSI"],
        loan_sorted["total_loan_created"].astype(int)
    )
]

loan_values = loan_sorted["total_os_loan_created"].astype(int).tolist()


# ======================================================
# 🔹 1. REJECT TOTAL (SLIK + SICD + BLACKLIST + SCORE + FAIL)
# ======================================================

reject_sorted = (
    gdf[["PROVINSI", "total_reject"]]
    .sort_values("total_reject", ascending=False)
    .head(10)
)

reject_labels = reject_sorted["PROVINSI"].tolist()
reject_values = reject_sorted["total_reject"].astype(int).tolist()


# ======================================================
# 🔹 2. USR REJECT (OS + NOA)
# ======================================================

# ✅ pakai kolom OS dari Excel
gdf["os_potensi_ktp"] = gdf["total_lepas_ktp_reject"].fillna(0)
gdf["os_potensi_usr"] = gdf["total_lepas_usr_reject"].fillna(0)
gdf["os_potensi"] = gdf["os_potensi_ktp"] + gdf["os_potensi_usr"]

usr_sorted = (
    gdf[["PROVINSI", "total_usr_reject", "os_potensi_usr"]]
    .sort_values("os_potensi_usr", ascending=False)
    .head(10)
)

usr_labels = [
    f"{prov} (NOA: {noa})"
    for prov, noa in zip(
        usr_sorted["PROVINSI"],
        usr_sorted["total_usr_reject"].astype(int)
    )
]

usr_values = usr_sorted["os_potensi_usr"].astype(int).tolist()


# ======================================================
# 🔹 3. KTP REJECT (OS + NOA)
# ======================================================

gdf["total_ktp_reject"] = gdf["total_ktp_reject"].fillna(0)
gdf["os_potensi_ktp"] = gdf["os_potensi_ktp"].fillna(0)

ktp_sorted = (
    gdf[["PROVINSI", "total_ktp_reject", "os_potensi_ktp"]]
    .sort_values("os_potensi_ktp", ascending=False)
    .head(10)
)

ktp_labels = [
    f"{prov} (NOA: {noa})"
    for prov, noa in zip(
        ktp_sorted["PROVINSI"],
        ktp_sorted["total_ktp_reject"].astype(int)
    )
]

ktp_values = ktp_sorted["os_potensi_ktp"].astype(int).tolist()

# === Reject Chart Top 10 ===
reject_sorted = (
    gdf[["PROVINSI", "total_reject"]]
    .sort_values("total_reject", ascending=False)
    .head(10)
)

reject_labels = reject_sorted["PROVINSI"].tolist()
reject_values = reject_sorted["total_reject"].astype(int).tolist()

# === 8. Build Map ===
m = folium.Map(location=[-2.5, 118], zoom_start=5, tiles="CartoDB positron")
Fullscreen().add_to(m)

colormap = cm.linear.YlGnBu_09.scale(
    float(gdf["total_loan_created"].min()),
    float(gdf["total_loan_created"].max())
)
colormap.caption = "Jumlah Pencairan (Loan Created)"

geojson = folium.GeoJson(
    gdf,
    style_function=lambda x: {
        "fillColor": colormap(x["properties"]["total_loan_created"]),
        "color": "black",
        "weight": 0.5,
        "fillOpacity": 0.7,
    },
    tooltip=folium.GeoJsonTooltip(
        fields=["PROVINSI", "total_loan_created", "total_reject",
                "os_potensi_ktp", "os_potensi_usr", "os_loan_created"],
        aliases=["Provinsi:", "Jumlah Loan Created:", "Jumlah Reject:",
                 "OS KTP Reject:", "OS USR Reject:", "OS Loan Created:"],
        localize=True
    )
).add_to(m)

colormap.add_to(m)

# === 9. Chart Container ===
chart_container = """
<div id="chartBox" style="
    position: fixed;
    top: 120px;
    right: 20px;
    z-index: 9999;
    width: 420px;
    background: rgba(255,255,255,0.95);
    padding: 12px;
    border-radius: 10px;
    box-shadow: 0 0 8px rgba(0,0,0,0.25);
    font-family: Arial;
">


    <!-- BUTTONS -->
    <div style="text-align:center; margin-bottom:10px;">
        <button id="btnLoan" onclick="showLoan()" style="padding:6px 10px; border-radius:6px; border:1px solid #0275d8; background:#0275d8; color:white; cursor:pointer;">
            Loan Created
        </button>

        <button id="btnReject" onclick="showReject()" style="padding:6px 10px; border-radius:6px; border:1px solid #d9534f; background:transparent; color:#d9534f; cursor:pointer;">
            Reject Total
        </button>

        <button id="btnUSR" onclick="showUSR()" style="padding:6px 10px; border-radius:6px; border:1px solid #ff8800; background:transparent; color:#ff8800; cursor:pointer;">
            USR Reject
        </button>

        <button id="btnKTP" onclick="showKTP()" style="padding:6px 10px; border-radius:6px; border:1px solid #6f42c1; background:transparent; color:#6f42c1; cursor:pointer;">
            KTP Reject
        </button>
    </div>

    <h4 id="chartTitle" style="margin:0; text-align:center; font-size:14px;">
        Loan Created — Top 10 OS
    </h4>

    <div style="height:300px; margin-top:10px;">
        <canvas id="chartReject"></canvas>
    </div>

</div>
"""
m.get_root().html.add_child(folium.Element(chart_container))

# === 10. Inject Chart Script ===
chart_script = f"""
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<script>

let chartObj = null;

// DATASET PYTHON → JAVASCRIPT
const loanLabels   = {json.dumps(loan_labels)};
const loanData     = {json.dumps(loan_values)};

const rejectLabels = {json.dumps(reject_labels)};
const rejectData   = {json.dumps(reject_values)};

const usrLabels    = {json.dumps(usr_labels)};
const usrData      = {json.dumps(usr_values)};

const ktpLabels    = {json.dumps(ktp_labels)};
const ktpData      = {json.dumps(ktp_values)};


function renderChart(labels, data, color) {{
    if (chartObj) chartObj.destroy();

    chartObj = new Chart(
        document.getElementById('chartReject').getContext('2d'),
        {{
            type: 'bar',
            data: {{
                labels: labels,
                datasets: [{{
                    label: 'Data',
                    data: data,
                    backgroundColor: color
                }}]
            }},
            options: {{
                indexAxis: 'y',
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }}
                }}
            }}
        }}
    );
}}

function activateButton(active, btns) {{
    btns.forEach(btn => {{
        const el = document.getElementById(btn);
        el.style.background = "transparent";
        el.style.color = el.style.borderColor;
    }});
    const act = document.getElementById(active);
    act.style.background = act.style.borderColor;
    act.style.color = "white";
}}

// === BUTTON BEHAVIOR ===

function showLoan() {{
    activateButton("btnLoan", ["btnLoan","btnReject","btnUSR","btnKTP"]);
    document.getElementById("chartTitle").innerText = "Loan Created — Top 10 OS (label NOA)";
    renderChart(loanLabels, loanData, "rgba(2,117,216,0.7)");
}}

function showReject() {{
    activateButton("btnReject", ["btnLoan","btnReject","btnUSR","btnKTP"]);
    document.getElementById("chartTitle").innerText = "Reject Total — Top 10 NOA";
    renderChart(rejectLabels, rejectData, "rgba(217,83,79,0.7)");
}}

function showUSR() {{
    activateButton("btnUSR", ["btnLoan","btnReject","btnUSR","btnKTP"]);
    document.getElementById("chartTitle").innerText = "USR Reject — Top 10 OS (label NOA)";
    renderChart(usrLabels, usrData, "rgba(255,136,0,0.7)");
}}

function showKTP() {{
    activateButton("btnKTP", ["btnLoan","btnReject","btnUSR","btnKTP"]);
    document.getElementById("chartTitle").innerText = "KTP Reject — Top 10 OS (label NOA)";
    renderChart(ktpLabels, ktpData, "rgba(111,66,193,0.7)");
}}

// DEFAULT = LOAN CREATED
showLoan();

</script>
"""

m.get_root().html.add_child(folium.Element(chart_script))


# === 11. Save File ===
m.save("index.html")

print("File generated successfully.")
