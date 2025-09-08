import matplotlib.pyplot as plt

def plot_map(gdf, column=None):
    ax = gdf.plot(column=column, cmap='viridis', edgecolor='black', legend=bool(column))
    ax.set_title("Peta Provinsi Indonesia")
    plt.axis("off")
    plt.show()
