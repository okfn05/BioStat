import pandas as pd
import numpy as np
import scipy.stats as stats

# Membaca data dari file Excel
file_path = 'Sistolik.xlsx'  # Ganti dengan path file Anda

# Membaca data dari Excel
data = pd.read_excel(file_path)

# Asumsi bahwa data memiliki kolom 'Sistolik_Masuk' dan 'Sistolik_2_Minggu'
sistolik_masuk = data['Sistolik_Masuk'].values
sistolik_2_minggu = data['Sistolik_2_Minggu'].values

# Menghitung selisih tekanan darah sistolik
selisih = sistolik_masuk - sistolik_2_minggu

# Menghitung nilai rata-rata dan standar deviasi dari selisih
rata_rata_selisih = np.mean(selisih)
std_dev_selisih = np.std(selisih, ddof=1)

# Melakukan uji t berpasangan
t_stat, p_value = stats.ttest_rel(sistolik_masuk, sistolik_2_minggu)

# Menyimpan output ke dalam DataFrame
output_data = pd.DataFrame({
    'Statistic': ['Rata-rata selisih', 'Standar deviasi selisih', 'Nilai t', 'Nilai p'],
    'Value': [rata_rata_selisih, std_dev_selisih, t_stat, p_value]
})

# Menentukan apakah ada perbedaan signifikan
alpha = 0.05
if p_value < alpha:
    significance = "Ada perbedaan signifikan dalam tekanan darah sistolik antara saat masuk dan setelah 2 minggu pelatnas."
else:
    significance = "Tidak ada perbedaan signifikan dalam tekanan darah sistolik antara saat masuk dan setelah 2 minggu pelatnas."

# Menambahkan hasil penentuan signifikansi ke DataFrame
output_data = pd.concat([output_data, pd.DataFrame({'Statistic': ['Signifikansi'], 'Value': [significance]})], ignore_index=True)

# Simpan output ke dalam file Excel
output_file_path = 'Sistolik_analisis.xlsx'
output_data.to_excel(output_file_path, index=False)

print(f"Output telah disimpan ke dalam file Excel dengan nama: {output_file_path}")
