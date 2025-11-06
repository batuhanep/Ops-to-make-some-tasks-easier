import pandas as pd
import numpy as np
import sys


INPUT_FILENAME = "koç ailem price index control.xlsx"

EXCEL_OUTPUT_FILENAME = "outlier_products.xlsx"

OUTLIER_THRESHOLD = 0.40

def clean_price(price_str):
    if isinstance(price_str, (int, float)):
        return price_str
    try:
        # Fiyatları string'e çevir, . (binlik) kaldır, , (ondalık) -> . (ondalık) yap
        cleaned_str = str(price_str).replace('.', '').replace(',', '.')
        return float(cleaned_str)
    except (ValueError, TypeError):
        return np.nan


def analyze_price_outliers_to_excel(file_path, output_path, threshold):
    try:
        df = pd.read_excel(file_path)

    except FileNotFoundError:
        print(f"HATA: '{file_path}' dosya kontrol.")
        return
    except ImportError:
        print("error.")
        return
    except Exception as e:
        print(f"error file reading")
        return

    df['checked_avg_price_numeric'] = df['checked_avg_price'].apply(clean_price)
    df['minprice_min_price_numeric'] = df['minprice_min_price'].apply(clean_price)

    df['max_price'] = df[['checked_avg_price_numeric', 'minprice_min_price_numeric']].max(axis=1)
    df['price_diff_ratio'] = np.where(
        (df['max_price'] > 0) & (df['checked_avg_price_numeric'].notna()) & (df['minprice_min_price_numeric'].notna()),
        np.abs(df['checked_avg_price_numeric'] - df['minprice_min_price_numeric']) / df['max_price'],
        0
    )

    outliers_df = df[df['price_diff_ratio'] > threshold].copy()

    relevant_columns = [
        'csku_id',
        'listing_title',
        'checked_retailer_code',
        'checked_customer_product_identifier',
        'checked_avg_price',
        'minprice_retailer_code',
        'minprice_customer_product_identifier',
        'minprice_min_price',
        'price_diff_ratio'
    ]

    existing_relevant_columns = [col for col in relevant_columns if col in outliers_df.columns]
    final_outliers = outliers_df[existing_relevant_columns].sort_values(by='price_diff_ratio', ascending=False)

    final_outliers['price_diff_ratio'] = final_outliers['price_diff_ratio'].round(4)

    try:
        final_outliers.to_excel(output_path, index=False, engine='openpyxl')

        print("-" * 30)
        print(f"done.")
        print(f"Toplam {len(df)} ürün incelendi.")
        print(f"Fiyat farkı %{threshold * 100:.0f}'den fazla olan {len(final_outliers)} adet aykırı ürün .")
        print(f"Sonuçlar '{output_path}' (Excel) dosyasına kaydedildi.")
        print("-" * 30)

    except ImportError:
        print("hata")

if __name__ == "__main__":
    analyze_price_outliers_to_excel(INPUT_FILENAME, EXCEL_OUTPUT_FILENAME, OUTLIER_THRESHOLD)