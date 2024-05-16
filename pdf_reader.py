import pdfplumber
import pandas as pd
# Путь к PDF-файлу
pdf_path = 'C:\\Users\\stud\\Desktop\\Parceprob2-master\\Datasheet\\IFSC-3232DB-01\\IFSC-3232DB-01.pdf'
table2 = []
# Открытие PDF-файла
with pdfplumber.open(pdf_path) as pdf:
    # Проходим по каждой странице PDF
    for page in pdf.pages:
        # Извлекаем таблицы с помощью метода extract_tables()
        tables = page.extract_tables()

        # Выводим извлеченные таблицы
        for table in tables:
            table2.extend(table[2:17])

df = pd.DataFrame(table2, columns=["Part Number", "Inductance at 0 A", "Inductace TOL.", "DCR", "Heat rating curent DC", "Saturation current DC", "SRF TYP." ])
df.to_excel("init.xlsx", index=False)

with pdfplumber.open(pdf_path) as pdf:
    # Переменная для хранения текста
    full_text = ""

    # Проходим по каждой странице PDF
    for page in pdf.pages:
        # Извлекаем текст с текущей страницы и добавляем его к общему тексту
        full_text += page.extract_text()

# Вывод или сохранение полного текста
lines = full_text.splitlines()

print(lines[2])

print(df)
