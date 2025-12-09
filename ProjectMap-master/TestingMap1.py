import folium
import geopandas as gpd
import pandas as pd
import branca.colormap as cm
from folium.plugins import Fullscreen
import os
from branca.element import Figure, Html, MacroElement
import json



# --- 1. Load shapefile ---
shapefile_path = r"C:/Users/User/Downloads/ProjectMap-master/BATAS PROVINSI DESEMBER 2019 DUKCAPIL/BATAS_PROVINSI_DESEMBER_2019_DUKCAPIL.shp"
gdf = gpd.read_file(shapefile_path)
gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.01, preserve_topology=True)

# --- 2. Load Excel ---
data_path = "data_pencairan1.xlsx"
nilai_df = pd.read_excel(data_path)

# Gabungkan shapefile dengan data
gdf = gdf.merge(nilai_df, how="left", on="PROVINSI")


# --- 3. Hitung total reject gabungan ---
gdf["total_reject"] = (
    gdf["total_reject_slik"].fillna(0) +
    gdf["total_reject_sicd_raya"].fillna(0) +
    gdf["total_reject_sicd_bri"].fillna(0) +
    gdf["total_reject_blacklist_company"].fillna(0) +
    gdf["total_reject_score_500"].fillna(0) +
    gdf["total_reject_failed_to_approve"].fillna(0)
)

# --- 3a. Hitung OS Potensi & Loan Created ---
#gdf["os_potensi"] = gdf["total_lepas_usr_reject"].fillna(0) + gdf["total_lepas_ktp_reject"].fillna(0)
#gdf["os_loan_created"] = gdf["total_os_loan_created"].fillna(0) if "total_os_loan_created" in gdf.columns else 0


# --- 3a. Hitung OS Potensi & Loan Created ---
#def format_rupiah(x):
#    try:
#        return "Rp {:,.0f}".format(x).replace(",", ".")
#   except:
#        return "Rp 0"

#gdf["os_potensi_rp"] = gdf["os_potensi"].apply(format_rupiah)
#gdf["os_loan_created_rp"] = gdf["os_loan_created"].apply(format_rupiah)

# --- 3a. Hitung OS Potensi (KTP & USR) ---
gdf["os_potensi_ktp"] = gdf["total_lepas_ktp_reject"].fillna(0)
gdf["os_potensi_usr"] = gdf["total_lepas_usr_reject"].fillna(0)

# Total os_potensi = KTP + USR
gdf["os_potensi"] = gdf["os_potensi_ktp"] + gdf["os_potensi_usr"]

# Loan created (kalau ada kolomnya)
gdf["os_loan_created"] = (
    gdf["total_os_loan_created"].fillna(0) if "total_os_loan_created" in gdf.columns else 0
)

# --- Format Rupiah ---
def format_rupiah(x):
    try:
        return "Rp {:,.0f}".format(x).replace(",", ".")
    except:
        return "Rp 0"

gdf["os_loan_created_rp"] = gdf["os_loan_created"].apply(format_rupiah)
gdf["os_potensi_ktp_rp"] = gdf["os_potensi_ktp"].apply(format_rupiah)
gdf["os_potensi_usr_rp"] = gdf["os_potensi_usr"].apply(format_rupiah)
gdf["os_potensi_rp"] = gdf["os_potensi"].apply(format_rupiah)


# --- 3b. Siapkan data untuk mini chart ---

# Pastikan NaN jadi 0 biar aman
gdf["total_loan_created"] = gdf["total_loan_created"].fillna(0)
gdf["total_reject"] = gdf["total_reject"].fillna(0)

# === Persiapan data untuk Mini Chart (Top 10) ===

# Pastikan tidak ada NaN
gdf["total_loan_created"] = gdf["total_loan_created"].fillna(0)
gdf["total_reject"] = gdf["total_reject"].fillna(0)

loan_sorted = (
    gdf[["PROVINSI", "total_loan_created"]]
    .dropna(subset=["PROVINSI"])
    .sort_values("total_loan_created", ascending=False)
    .head(10)   # <<< hanya top 10
)

reject_sorted = (
    gdf[["PROVINSI", "total_reject"]]
    .dropna(subset=["PROVINSI"])
    .sort_values("total_reject", ascending=False)
    .head(10)   # <<< hanya top 10
)

loan_labels = loan_sorted["PROVINSI"].tolist()
loan_values = loan_sorted["total_loan_created"].astype(int).tolist()

reject_labels = reject_sorted["PROVINSI"].tolist()
reject_values = reject_sorted["total_reject"].astype(int).tolist()


# --- 4. Inisialisasi peta ---
m = folium.Map(location=[-2.5, 118], zoom_start=5, tiles="CartoDB positron")
Fullscreen().add_to(m)

# --- 5. Colormap Loan ---
colormap = cm.linear.YlGnBu_09.scale(gdf["total_loan_created"].min(), gdf["total_loan_created"].max())
colormap.caption = "Jumlah Pencairan (Loan Created)"

# --- 6. GeoJson Layer ---
geojson = folium.GeoJson(
    gdf,
    style_function=lambda x: {
        "fillColor": colormap(x["properties"]["total_loan_created"]) if x["properties"]["total_loan_created"] is not None else "#d3d3d3",
        "color": "black",
        "weight": 0.5,
        "fillOpacity": 0.7,
    },
    tooltip=folium.GeoJsonTooltip(
        fields=["PROVINSI", "total_loan_created", "total_reject", "os_potensi_ktp","os_potensi_usr", "os_loan_created"],
        aliases=["Provinsi:", "Jumlah Loan Created:", "Jumlah Reject:", "OS Potensi KTP Reject:","OS Potensi USR Reject:", "OS Loan Created:"],
        localize=True
    ),
   popup = folium.GeoJsonPopup(
    fields=[
        "PROVINSI",
        "total_reject_slik",
        "total_reject_sicd_raya",
        "total_reject_blacklist_company",
        "total_ktp_reject",
        "total_usr_reject"
    ],
    aliases=[
        "Provinsi:",
        "Reject SLIK:",
        "Reject SICD RAYA:",
        "Reject Blacklist Company:",
        "Reject KTP:",
        "Reject USR:"
    ],
    localize=True,
    labels=True,
    max_width=400
)
).add_to(m)

# --- 7. Tambahkan colormap ---
colormap.add_to(m)

# --- Tambahkan Judul ---
title_html = """
     <h3 align="center" style="font-size:20px">
         <b>DATA LOAN CREATED DAN REJECT NOVEMBER 2025</b>
     </h3>
     """
m.get_root().html.add_child(folium.Element(title_html))

# --- 10. Tambahkan container untuk 2 mini chart (pojok kanan atas) ---

chart_container = """
<div id="chartBox" style="
    position: fixed;
    top: 120px;
    right: 20px;
    z-index: 9999;
    width: 380px;
    background: rgba(255,255,255,0.95);
    padding: 12px;
    border-radius: 10px;
    box-shadow: 0 0 8px rgba(0,0,0,0.25);
    font-family: Arial;
">

    <!-- TOGGLE BUTTONS -->
    <div style="text-align:center; margin-bottom:10px;">
        <button id="btnLoan" onclick="showLoan()" style="
            padding:6px 12px;
            border-radius:6px;
            border:1px solid #0275d8;
            background:#0275d8;
            color:white;
            cursor:pointer;
        ">Loan Created</button>

        <button id="btnReject" onclick="showReject()" style="
            padding:6px 12px;
            border-radius:6px;
            border:1px solid #d9534f;
            background:transparent;
            color:#d9534f;
            cursor:pointer;
        ">Reject</button>
    </div>

    <!-- LOAN CHART -->
    <div id="loanSection" style="height:200px;">
        <canvas id="loanChart"></canvas>
    </div>

    <!-- REJECT CHART -->
    <div id="rejectSection" style="height:200px; display:none;">
        <canvas id="rejectChart"></canvas>
    </div>

</div>
"""

m.get_root().html.add_child(folium.Element(chart_container))

# --- 11. Script Chart.js untuk menggambar 2 mini chart ---

chart_script = f"""
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<script>
function showLoan() {{
    document.getElementById("loanSection").style.display = "block";
    document.getElementById("rejectSection").style.display = "none";

    // aktifkan tombol
    document.getElementById("btnLoan").style.background = "#0275d8";
    document.getElementById("btnLoan").style.color = "white";

    document.getElementById("btnReject").style.background = "transparent";
    document.getElementById("btnReject").style.color = "#d9534f";
}}

function showReject() {{
    document.getElementById("loanSection").style.display = "none";
    document.getElementById("rejectSection").style.display = "block";

    // aktifkan tombol reject
    document.getElementById("btnReject").style.background = "#d9534f";
    document.getElementById("btnReject").style.color = "white";

    document.getElementById("btnLoan").style.background = "transparent";
    document.getElementById("btnLoan").style.color = "#0275d8";
}}

const loanLabels = {json.dumps(loan_labels)};
const loanData = {json.dumps(loan_values)};
const rejectLabels = {json.dumps(reject_labels)};
const rejectData = {json.dumps(reject_values)};

// CHART LOAN
new Chart(document.getElementById('loanChart').getContext('2d'), {{
    type: 'bar',
    data: {{
        labels: loanLabels,
        datasets: [{{
            label: 'Loan Created',
            data: loanData,
            backgroundColor: 'rgba(2,117,216,0.7)',
        }}]
    }},
    options: {{
        indexAxis: 'y',
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }}
    }}
}});

// CHART REJECT
new Chart(document.getElementById('rejectChart').getContext('2d'), {{
    type: 'bar',
    data: {{
        labels: rejectLabels,
        datasets: [{{
            label: 'Reject',
            data: rejectData,
            backgroundColor: 'rgba(217,83,79,0.7)',
        }}]
    }},
    options: {{
        indexAxis: 'y',
        maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }}
    }}
}});
</script>
"""

m.get_root().html.add_child(folium.Element(chart_script))



# --- 8. Simpan HTML ---
m.save("index.html")

# --- 9. Cek ukuran file ---
file_path = "index.html"
size_mb = os.path.getsize(file_path) / (1024*1024)
print(f"Ukuran file HTML: {size_mb:.2f} MB")
