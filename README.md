import pandas as pd
import openpyxl
import random
from tkinter import filedialog  # Dosya seçim penceresi için

def random_data_to_excel(output_file, header_categories, num_rows_to_extract):
    # Excel dosyasını seçin
    input_file = filedialog.askopenfilename(title="Excel Dosyasını Seçin", filetypes=[("Excel Dosyası", "*.xlsx")])

    # Excel dosyasını oku
    df = pd.read_excel(input_file)

    # Rastgele satırları seç
    selected_rows = df.sample(n=num_rows_to_extract)

    # Seçilen satırların sadece belirtilen başlık kategorilerine ait sütunlarını al
    selected_columns = [col for col in df.columns if any(category in col for category in header_categories)]
    selected_df = selected_rows[selected_columns]

    # Kategori değişkeni oluştur
    selected_df['Category'] = random.choices(header_categories, k=num_rows_to_extract)

    # Kategoriye göre grupla ve alt alta yerleştir
    grouped_df = selected_df.groupby('Category', as_index=False).apply(lambda x: x.reset_index(drop=True))

    # Excel dosyasına yaz
    grouped_df.to_excel(output_file, index=False)

# Kullanım örneği
output_excel = "cikis.xlsx"  # Çıkış Excel dosyasının adı
header_categories = ["Category1", "Category2", "Category3"]  # Başlık kategorilerinin listesi
num_rows_to_extract =3 # Rastgele kaç satır almak istediğinizi belirtin

random_data_to_excel(output_excel, header_categories, num_rows_to_extract)
