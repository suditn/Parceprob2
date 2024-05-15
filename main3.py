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
url = 'https://www.vishay.com/en/inductors/'
save_path = str(Path(__file__).parent.resolve())
img_small_save_path = os.path.join(save_path, "image", "small_inductors")
datasheet_save_path = os.path.join(save_path, "Datasheet")
headers = {'User-Agent': "scrapping_script/1.0"}

# Создание папок
def create_directories():
    Path(img_small_save_path).mkdir(parents=True, exist_ok=True)
    Path(datasheet_save_path).mkdir(parents=True, exist_ok=True)

# Выгрузка полной таблицы на страницу
def get_web(u):
    driver.get(u)
    logging.info(f'Открыта страница: {u}')
    create_directories()
    wait = WebDriverWait(driver, 10)
    option = driver.find_element('xpath', '//label/select/option[1]')
    option_max = driver.find_element('xpath', '//label/select/option[3]')
    try:
        max_entries = driver.find_element('xpath', '/html/body/div[1]/div/div[2]/div[2]/div[2]/div/div/div/div/div[3]/div[1]/div').text
    except UnboundLocalError:
        option_max.click()
    finally:
        driver.execute_script('arguments[0].value = arguments[1]', option, pattern.findall(max_entries)[5])
        option.click()
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
                return True  # Возвращаем успешное завершение загрузки
        except Exception as e:
            logging.error(f'Ошибка при загрузке файла {url}: {e}')
            logging.info(f'Повторная попытка загрузки файла {url}...')
            continue
    return False  # Возвращаем неудачное завершение загрузки после всех попыток

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
                return True  # Возвращаем успешное завершение загрузки
        except Exception as e:
            logging.error(f'Ошибка при загрузке изображения {url}: {e}')
            logging.info(f'Повторная попытка загрузки изображения {url}...')
            continue
    return False  # Возвращаем неудачное завершение загрузки после всех попыток

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
                    if b.endswith('_3dmodel.zip'):
                        return download_file_with_retry('https://www.vishay.com/' + b, file_3d_path, headers)
            return False  # Не удалось найти 3D модель
    except requests.exceptions.RequestException as e:
        logging.error(f'Ошибка при получении 3D модели для продукта {img_alt}: {e}')
        return False  # Не удалось получить 3D модель
    return False  # Не удалось получить 3D модель

# Функция для обработки HTML и запуска параллельной загрузки.
def process_html(html_source):
    soup = BeautifulSoup(html_source, "lxml")
    table = soup.find('table', {'id': 'poc'})
    images = table.findAll('img')
    columns = [i.get_text(strip=True) for i in table.find_all("th")]
    data = [[td.get_text(strip=True) for td in tr.find_all("td")] for tr in table.find("tbody").find_all("tr")]

    df = pd.DataFrame(data, columns=columns)
    img_src = []
    datasheet_src = []
    file_3d_src = []
    previous_img_src = ''
    previous_datasheet_src = ''
    file_3d_exsist = False
    i = 0
    imgpr = ''

    for img in images:
        series = df['Series▲▼'][i]
        if img['src'].split('/')[-2] == 'pt-small':
            img_filename = img['alt'] + '.png'
            img_path = os.path.join(img_small_save_path, img_filename)
            img_src.append(img_path)
            if previous_img_src != img['src'] and img['alt'] != "Datasheet":
                download_image_with_retry('https://www.vishay.com/' + img['src'], img_path, headers)
                previous_img_src = img['src']

            datasheet_filename = series + '.pdf'
            file_3d_name = series+'.zip'
            datasheet_path = os.path.join(datasheet_save_path, series, datasheet_filename)
            file_3d_path = os.path.join(datasheet_save_path,series,file_3d_name)
            datasheet_src.append(datasheet_path)

            if previous_datasheet_src != series and img['alt'] != "Datasheet":
                download_file_with_retry('https://www.vishay.com/doc?' + img['alt'], datasheet_path, headers)
                if download_3d_model_with_retry(img['alt'], file_3d_path) == True:
                    file_3d_exsist == True


            if file_3d_exsist==True:
                file_3d_src.append(file_3d_path)
            else:
                file_3d_src.append('3Д модели нету')

            if series != previous_datasheet_src:
                file_3d_exsist = False

            imgpr = img['alt']

            previous_datasheet_src = series
            i += 1


    return df, img_src, datasheet_src, file_3d_src

# Функция для сохранения данных в Excel.
def save_to_excel(df, img_src, datasheet_src, file_3d_src, save_path, url):
    excel_path = os.path.join(save_path, url.split('/')[-2] + '.xlsx')
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df_img = pd.DataFrame(img_src, columns=['Product Image'])
        df_datasheet = pd.DataFrame(datasheet_src, columns=['Datasheet'])
        df_file_3d = pd.DataFrame(file_3d_src, columns=['3D Models'])
        df_final = df.join(df_img, lsuffix='_left', rsuffix='_right').join(df_datasheet, rsuffix='_datasheet').join(df_file_3d, rsuffix='_datasheet')
        df_final.to_excel(writer, index=False, sheet_name='Inductors')
        worksheet = writer.sheets['Inductors']
        worksheet.autofit()

# Остальной код остается неизменным

try:
    web_source = get_web(url)
    df, img_src, datasheet_src, file_3d_src = process_html(web_source)
    save_to_excel(df, img_src, datasheet_src, file_3d_src, save_path, url)
    logging.info('Данные успешно сохранены.')
finally:
    driver.quit()
