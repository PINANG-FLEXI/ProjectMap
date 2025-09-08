import folium
import geopandas as gpd
import pandas as pd
import branca.colormap as cm
from folium.plugins import Fullscreen
import os

# --- 1. Load data & Sederhanakan Geometri ---
shapefile_path = "C:\\Users\\User\\Downloads\\ProjectMap-master\\BATAS PROVINSI DESEMBER 2019 DUKCAPIL\\BATAS_PROVINSI_DESEMBER_2019_DUKCAPIL.shp"
gdf = gpd.read_file(shapefile_path)

# PENTING: Sederhanakan data untuk mengurangi ukuran file
# Coba toleransi yang lebih tinggi jika file masih terlalu besar
gdf['geometry'] = gdf['geometry'].simplify(tolerance=0.01, preserve_topology=True)

nilai_df1 = pd.read_excel("data_pencairan1.xlsx")
nilai_df2 = pd.read_excel("data_pencairan2.xlsx")

gdf1 = gdf.merge(nilai_df1, how="left", on="PROVINSI")
gdf2 = gdf.merge(nilai_df2, how="left", on="PROVINSI")

# --- 2. Simpan GeoDataFrame menjadi file GeoJSON terpisah ---
# Pastikan nama file ini konsisten
gdf1.to_file("data1.geojson", driver="GeoJSON")
gdf2.to_file("data2.geojson", driver="GeoJSON")

# --- 3. Inisialisasi peta ---
m = folium.Map(location=[-2.5, 118], zoom_start=5, tiles="CartoDB positron")
Fullscreen().add_to(m) # Anda lupa menambahkan Fullscreen() di kode terakhir Anda.

# --- 4. Buat colormap & Tambahkan layer DARI FILE GEOJSON ---
all_vals = pd.concat([gdf1["nilai"], gdf2["nilai"]])
colormap = cm.linear.YlGnBu_09.scale(all_vals.min(), all_vals.max())
colormap.caption = "Jumlah Pencairan"

folium.GeoJson(
    "data1.geojson", # Mengacu pada file GeoJSON
    name="Jumlah Pencairan Tanggal 2025-07-01 s/d 2025-07-14",
    style_function=lambda x: {
        "fillColor": colormap(x["properties"]["nilai"]) if x["properties"]["nilai"] is not None else "#d3d3d3",
        "color": "black",
        "weight": 0.5,
        "fillOpacity": 0.7,
    },
    tooltip=folium.GeoJsonTooltip(
        fields=["PROVINSI", "nilai"],
        aliases=["Provinsi:", "Jumlah:"],
        localize=True
    )
).add_to(m)

folium.GeoJson(
    "data2.geojson", # Mengacu pada file GeoJSON
    name="Jumlah Reject Slik 2025-07-01 s/d 2025-07-14",
    style_function=lambda x: {
        "fillColor": colormap(x["properties"]["nilai"]) if x["properties"]["nilai"] is not None else "#d3d3d3",
        "color": "black",
        "weight": 0.5,
        "fillOpacity": 0.7,
    },
    tooltip=folium.GeoJsonTooltip(
        fields=["PROVINSI", "nilai"],
        aliases=["Provinsi:", "Jumlah:"],
        localize=True
    )
).add_to(m)

# --- 5. Tambahkan kontrol layer & colormap ---
colormap.add_to(m)
folium.LayerControl(collapsed=False).add_to(m)

# --- 6. Simpan HTML ---
m.save("index.html") # Pastikan Anda menyimpannya sebagai index.html

# Opsional: Periksa ukuran file
file_path = "index.html"
size_mb = os.path.getsize(file_path) / (1024*1024)
print(f"Ukuran file HTML: {size_mb:.2f} MB")