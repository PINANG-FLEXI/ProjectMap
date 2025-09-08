import folium
import geopandas as gpd
import pandas as pd
import branca.colormap as cm
from folium.plugins import Fullscreen
import os

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
        fields=["PROVINSI", "total_loan_created", "total_reject"],
        aliases=["Provinsi:", "Jumlah Loan:", "Jumlah Reject:"],
        localize=True
    ),
    popup=folium.GeoJsonPopup(
        fields=[
            "PROVINSI",
            "total_loan_created",
            "total_reject_slik",
            "total_reject_sicd_raya",
            "total_reject_sicd_bri",
            "total_reject_blacklist_company",
            "total_reject_score_500",
            "total_reject_failed_to_approve"
        ],
        aliases=[
            "Provinsi:",
            "Total Loan:",
            "Reject SLIK:",
            "Reject SICD RAYA:",
            "Reject SICD BRI:",
            "Reject BLACKLIST COMPANY:",
            "Reject SCORE BELOW 500:",
            "Reject Failed Approve:"
        ],
        localize=True,
        labels=True,
        max_width=400
    )
).add_to(m)

# --- 7. Tambahkan colormap ---
colormap.add_to(m)

# --- 8. Simpan HTML ---
m.save("index.html")

# --- 9. Cek ukuran file ---
file_path = "index.html"
size_mb = os.path.getsize(file_path) / (1024*1024)
print(f"Ukuran file HTML: {size_mb:.2f} MB")
