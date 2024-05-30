import os
import shutil
import time
from pathlib import Path
import logging
import requests
from bs4 import BeautifulSoup
import pandas as pd

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.options import Options

# Логи
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Поиск слова
import re
pattern = re.compile(r'\w+')

# Настройка браузера ссылок и путей файлов
options = Options()
options.headless = True
driver = webdriver.Edge(options=options)
url = 'https://www.vishay.com/en'
save_path = str(Path(__file__).parent.resolve())
print(save_path)
img_small_save_path = os.path.join(save_path, "image")
print(img_small_save_path)
datasheet_save_path = os.path.join(save_path, "Datasheet")
headers = {'User-Agent': "scrapping_script/1.0"}
#"/diodes/emi-filter/", "/diodes/zener-stabilizers/", "/diodes/switching/" , "/diodes/ss-schottky/","/diodes/standard-recovery/", "/diodes/ultrafast-recovery/",
#"/diodes/silicon-carbide/",
sl_use = [  "/diodes/schottky/", "/diodes/bridge/", "/diodes/med-high-diodes/"]

# Создание папок
def create_directories(sl):
    Path(img_small_save_path, sl).mkdir(parents=True, exist_ok=True)
    Path(datasheet_save_path, sl).mkdir(parents=True, exist_ok=True)

# Выгрузка полной таблицы на страницу
def get_web(u, sl):
    driver.get(u)
    print (u)
    logging.info(f'Открыта страница: {u}')
    create_directories(sl)
    wait = WebDriverWait(driver, 10)
    option = driver.find_element('xpath', '//label/select/option[1]')
    option_max = driver.find_element('xpath', '//label/select/option[3]')
    try:
        max_entries = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div').text
        driver.execute_script('arguments[0].value = arguments[1]', option, pattern.findall(max_entries)[5])
        option.click()
    except UnboundLocalError:
        option_max.click()
    finally:
        return driver.page_source

# Функция для скачивания файлов с повторными попытками
def download_file_with_retry(url, path, headers=None):
    time.sleep(0.3)
    retries = 3
    for _ in range(retries):
        try:
            with requests.get(url, headers=headers, stream=True) as response:
                response.raise_for_status()
                file_dir = os.path.dirname(path)
                Path(file_dir).mkdir(parents=True, exist_ok=True)
                if not os.path.exists(path):
                    with open(path, 'wb') as out_file:
                        shutil.copyfileobj(response.raw, out_file)
                    logging.info(f'Файл {os.path.basename(path)} успешно загружен и сохранен.')
                else:
                    logging.info(f'Файл {os.path.basename(path)} уже существует.')
                return True
        except Exception as e:
            logging.error(f'Ошибка при загрузке файла {url}: {e}')
            logging.info(f'Повторная попытка загрузки файла {url}...')
            continue
    return False

# Функция для скачивания изображений с повторными попытками
def download_image_with_retry(url, path, headers=None):
    time.sleep(0.3)
    retries = 3
    for _ in range(retries):
        try:
            with requests.get(url, headers=headers, stream=True) as response:
                response.raise_for_status()
                file_dir = os.path.dirname(path)
                Path(file_dir).mkdir(parents=True, exist_ok=True)
                if not os.path.exists(path):
                    with open(path, 'wb') as out_file:
                        shutil.copyfileobj(response.raw, out_file)
                    logging.info(f'Изображение {os.path.basename(path)} успешно загружено и сохранено.')
                else:
                    logging.info(f'Изображение {os.path.basename(path)} уже существует.')
                return True
        except Exception as e:
            logging.error(f'Ошибка при загрузке изображения {url}: {e}')
            logging.info(f'Повторная попытка загрузки изображения {url}...')
            continue
    return False

def download_3d_model_with_retry(img_alt, file_3d_path):
    try:
        with requests.get('https://www.vishay.com/en/product/' + img_alt + '/tab/designtools-ppg/', stream=True, headers=headers, timeout=10) as response:
            response.raise_for_status()
            if response.status_code == 200:
                soupp = BeautifulSoup(response.content, "lxml")
                file_3d_cont = []
                for a in soupp.findAll('a', href=True):
                    file_3d_cont.append(a['href'])

                for b in file_3d_cont:
                    if b.endswith('.zip') or b.endswith('.txt'):
                        return download_file_with_retry('https://www.vishay.com/' + b, file_3d_path, headers)
            return False
    except requests.exceptions.RequestException as e:
        logging.error(f'Ошибка при получении 3D модели для продукта {img_alt}: {e}')
        return False
    return False

# Функция для обработки HTML и запуска параллельной загрузки.
def process_html(html_source, sl):
    soup = BeautifulSoup(html_source, "lxml")
    table = soup.find('table', {'id': 'poc'})
    images = table.findAll('img')
    columns = [i.get_text(strip=True) for i in table.find_all("th")]
    data = [[td.get_text(strip=True) for td in tr.find_all("td")] for tr in table.find("tbody").find_all("tr")]
    df = pd.DataFrame(data, columns=columns)
    print(columns)
    img_src = []
    datasheet_src = []
    file_3d_src = []
    previous_img_src = ''
    previous_datasheet_src = ''
    file_3d_exsist = False
    i = 0
    imgpr = ''

    for img in images:
        try:
            series = df['Series▲▼'][i]
        except KeyError:
            try:
                series = df['Part Number▲▼'][i]
            except KeyError:
                series = df['Part Number▲▼(all)VEMI45AA-HNHVEMI45AB-HNHVEMI45AC-HNHVEMI65AA-HCIVEMI65AB-HCIVEMI65AC-HCIVEMI85AA-HGKVEMI85AB-HGKVEMI85AC-HGK'][i]
        if img['src'].split('/')[-2] == 'pt-small':
            img_filename = img['alt'] + '.png'
            img_src.append("image"+sl+img_filename)
            img_path = os.path.join(img_small_save_path+sl, img_filename)

            if previous_img_src != img['src'] and img['alt'] != "Datasheet":
                download_image_with_retry('https://www.vishay.com/' + img['src'], img_path, headers)
                previous_img_src = img['src']

            datasheet_filename = series + '.pdf'
            file_3d_name = series+'.zip'
            datasheet_path = os.path.join(datasheet_save_path + sl, series, datasheet_filename)
            file_3d_path = os.path.join(datasheet_save_path + sl, series, file_3d_name)

            if previous_datasheet_src != series and img['alt'] != "Datasheet":
                download_file_with_retry('https://www.vishay.com/doc?' + img['alt'], datasheet_path, headers)
                download_3d_model_with_retry(img['alt'], file_3d_path)

            imgpr = img['alt']
            previous_datasheet_src = series
            i += 1

    return df, img_src

# Функция для сохранения данных в Excel.
def save_to_excel(df, img_src, save_path, url):
    excel_path = os.path.join(save_path, url.split('/')[-2] + '.xlsx')
    print(excel_path)
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df_img = pd.DataFrame(img_src, columns=['Image path'])
        df_final = df.join(df_img, rsuffix='_datasheet')
        df_final.to_excel(writer, index=False, sheet_name=url.split('/')[-2])
        worksheet = writer.sheets[url.split('/')[-2]]
        worksheet.autofit()

# Функция для сохранения данных в CSV.
def save_to_csv(df, img_src, save_path, url):
    csv_path = os.path.join(save_path, url.split('/')[-2] + '.csv')
    print(csv_path)
    df_img = pd.DataFrame(img_src, columns=['Image path'])
    df_final = df.join(df_img, rsuffix='_datasheet')
    df_final.to_csv(csv_path, index=False)

try:
    for sl in sl_use:
        web_source = get_web(url+sl, sl)
        df, img_src = process_html(web_source, sl)
        save_to_excel(df, img_src, save_path, url+sl)
        save_to_csv(df, img_src, save_path, url+sl)
        logging.info('Данные успешно сохранены.')
finally:
    print("Всё")
