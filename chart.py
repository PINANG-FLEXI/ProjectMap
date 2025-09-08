import pandas as pd
import matplotlib.pyplot as plt

# Contoh data (bisa diganti nanti dari Excel)
data = {
    "state_desc": ["BALI", "BANTEN", "JAWA BARAT", "JAWA TENGAH"],
    "total_nasabah": [15, 119, 235, 144]
}

df = pd.DataFrame(data)

# Plot bar chart
plt.figure(figsize=(10,6))
bars = plt.bar(df["state_desc"], df["total_nasabah"], color="royalblue")

# Tambahkan angka di atas batang
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval + 5, str(yval),
             ha='center', va='bottom', fontsize=9)

plt.title("Total Pencairan per Wilayah")
plt.xlabel("Wilayah")
plt.ylabel("Total Nasabah")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()