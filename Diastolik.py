import pandas as pd
from scipy import stats

# Membaca data dari file Excel
data = pd.read_excel('Diastolik.xlsx')

# Memisahkan data berdasarkan baris
male_data = data.iloc[:14].copy()
female_data = data.iloc[14:30].copy()

# Menghitung perubahan tekanan darah diastolik
male_data['change'] = male_data['diastolic_after'] - male_data['diastolic_before']
female_data['change'] = female_data['diastolic_after'] - female_data['diastolic_before']

# Melakukan uji t berpasangan untuk masing-masing gender
t_stat_male, p_value_male = stats.ttest_rel(male_data['diastolic_before'], male_data['diastolic_after'])
t_stat_female, p_value_female = stats.ttest_rel(female_data['diastolic_before'], female_data['diastolic_after'])

# Melakukan uji t independen untuk perubahan antara dua gender
t_stat_change, p_value_change = stats.ttest_ind(male_data['change'], female_data['change'])

# Menentukan apakah ada perbedaan yang signifikan dalam perubahan tekanan darah diastolik
alpha = 0.05
if p_value_change < alpha:
    significance = "Terdapat perbedaan signifikan dalam perubahan tekanan darah diastolik antara atlet laki-laki dan perempuan."
else:
    significance = "Tidak terdapat perbedaan signifikan dalam perubahan tekanan darah diastolik antara atlet laki-laki dan perempuan."

# Membuat DataFrame untuk output
output_data = {
    "Statistic": ["t-statistic (male)", "p-value (male)", "t-statistic (female)", "p-value (female)",
                  "t-statistic (change)", "p-value (change)", "Significance"],
    "Value": [t_stat_male, p_value_male, t_stat_female, p_value_female,
              t_stat_change, p_value_change, significance]
}
output_df = pd.DataFrame(output_data)

# Simpan output ke file Excel
output_df.to_excel('output_statistik_diastolik.xlsx', index=False)
