import geopandas as gpd
import os

def load_shapefile():
    path = os.path.join("..", "data", "BATAS_PROVINSI_DESEMBER_2019_DUKCAPIL")
    for file in os.listdir(path):
        if file.endswith(".shp"):
            shp_path = os.path.join(path, file)
            return gpd.read_file(shp_path)
    raise FileNotFoundError("Shapefile tidak ditemukan.")
