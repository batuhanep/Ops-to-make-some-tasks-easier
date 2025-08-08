import pandas as pd

input_file = "amzfill.xlsx"  # ← Buraya kendi dosya adını yaz
output_file = "amzfillfinal_doldurulmus.xlsx"

df = pd.read_excel(input_file)


columns_to_fill = ['Brand', 'Barcode', 'Title']
df[columns_to_fill] = df[columns_to_fill].fillna(method='ffill')

df.to_excel(output_file, index=False)

print(f"Tamamlandı! Doldurulmuş dosya: {output_file}")
