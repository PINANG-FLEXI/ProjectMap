from load_shapefile import load_shapefile
from load_excel import load_excel
from plot_map import plot_map

def main():
    gdf = load_shapefile()
    df_excel = load_excel()

    # contoh join (pastikan kolom "PROVINSI" ada di kedua file)
    gdf_merged = gdf.merge(df_excel, on="PROVINSI", how="left")

    plot_map(gdf_merged, column="nilai")  # ubah sesuai nama kolom dari excel

if __name__ == "__main__":
    main()
