import folium
import geopandas as gpd
import pandas as pd
import branca.colormap as cm
from folium.plugins import Fullscreen
import os
from branca.element import Figure, Html, MacroElement


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

gdf["os_potensi_ktp_rp"] = gdf["os_potensi_ktp"].apply(format_rupiah)
gdf["os_potensi_usr_rp"] = gdf["os_potensi_usr"].apply(format_rupiah)
gdf["os_potensi_rp"] = gdf["os_potensi"].apply(format_rupiah)
gdf["os_loan_created_rp"] = gdf["os_loan_created"].apply(format_rupiah)

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
         <b>DATA LOAN_CREATED DAN REJECT AGUSTUS 2025</b>
     </h3>
     """
m.get_root().html.add_child(folium.Element(title_html))

# --- 8. Simpan HTML ---
m.save("index.html")

# --- 9. Cek ukuran file ---
file_path = "index.html"
size_mb = os.path.getsize(file_path) / (1024*1024)
print(f"Ukuran file HTML: {size_mb:.2f} MB")
