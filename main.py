#Tutaj podaj sieceżkę do pliku z danymi do logowania do Google Sheets
path = "D:\\PythonScripts\\SheetsProject\\credentials.json"

import gspread
from bs4 import BeautifulSoup
from google.oauth2.service_account import Credentials
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Funkcja pomocnicza do konwersji numeru kolumny na litery Excela
def column_number_to_letter(n):
    result = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

# zwraca drivera z odpowiednimi ustawieniami
def driver_settings():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# zwraca listę 10 elementów jeśli jest z PL
def scrap_financial_data(ticker):
    tir = str(ticker)
    if tir[-1] == 'S' and tir[-2] == 'U' and tir[-3] == '.':
        print("TAK! BINOGO, AMERYKA!")
    strona = "https://www.biznesradar.pl/notowania/" + ticker
    driver = driver_settings()
    data_exchange = []  
    try:
        driver.get(strona)
        print("Udało sie połączyć")
        time.sleep(2)
    except:
        print(f"Nie udało się otworzyć strony - {ticker}")
        driver.quit()
    try:
        element = driver.find_elements(By.CLASS_NAME, 'q_ch_act')
        data_exchange.append(element[0].text)
        print(data_exchange[0])
        nowa = []
        nowszaXD = []
        nowa.append(f'kurs - {data_exchange[0]}')
        elements = driver.find_elements(By.TAG_NAME, 'td')
        for i in elements:
            nowszaXD.append(i.get_attribute('innerHTML').strip())
        teksthtml = ''.join(nowszaXD)
        soup = BeautifulSoup(teksthtml, 'html.parser')
        values = soup.find_all('span')
        l=0
        if tir[-1] == 'S' and tir[-2] == 'U' and tir[-3] == '.':
            print("TAK! BINOGO, AMERYKA!")
            nowa.append(f'Zmiana 1d: {values[8-2].text}')
            print(values[8-2].text)
            nowa.append(f'Zmiana 7d: {values[10-1].text}')
            print(values[10-1].text)
            nowa.append(f'Zmiana 1m: {values[12-1].text}')
            print(values[12-1].text)
            nowa.append(f'Zmiana 3m: {values[14-1].text}')
            print(values[14-1].text)
            nowa.append(f'Zmiana 12m: {values[18-1].text}')
            print(values[18-1].text)
            nowa.append("-"*30)
          
        else:
            nowa.append(f'Zmiana 1d: {values[8].text}')
            print(values[8].text)
            nowa.append(f'Zmiana 7d: {values[11].text}')
            print(values[10-1].text)
            nowa.append(f'Zmiana 1m: {values[13].text}')
            print(values[12-1].text)
            nowa.append(f'Zmiana 3m: {values[15].text}')
            print(values[14-1].text)
            nowa.append(f'Zmiana 12m: {values[19].text}')
            print(values[18-2-1].text)
            nowa.append(f'Przychody ze sprzedaży r/r: {values[20-1-1].text}')
            print(values[20-1-1].text)
            nowa.append(f'EBIT r/r: {values[25-2-1].text}')
            print(values[25-2-1].text)
            nowa.append(f'Zysk neeto: r/r: {values[40-1-1].text}')
            print(values[40-1-1].text)
            nowa.append(f'C/Z: {values[68-2].text}')
            print(values[68-1].text)
    except:
        nowa.append("Nie udało się znaleźć elementów")
        driver.quit()
    print("-"*30)
    print(len(nowa))
    return nowa

# zwra listę 3 elementów: link, tytuł, treść dla wszstkich wiadmości
def scrap_news(ticker):
    tir = str(ticker)
    if tir[-1] == 'S' and tir[-2] == 'U' and tir[-3] == '.':
        print("TAK! BINOGO, AMERYKA!")
    strona = "https://www.biznesradar.pl/wiadomosci/" + ticker
    driver = driver_settings()
    try:
        driver.get(strona)
        print("Udało sie połączyć")
        time.sleep(2)
    except:
        print(f"Nie udało się otworzyć strony - {ticker}")
        driver.quit()

    try:
        element = driver.find_element(By.ID, 'news-radar-body')
        content = element.get_attribute('innerHTML')
        teksthtml = ''.join(content)
        soup = BeautifulSoup(teksthtml, 'html.parser')
        news_list = []
        news_items = soup.find_all('div', class_='record record-type-NEWS')
    except Exception as e:
        print(f"Nie udało się znaleźć elementów - {e}")
        driver.quit()
    for item in news_items:
        link = item.find('div', class_='record-header').find('a')['href']
        title = item.find('div', class_='record-header').find('a').text.strip()
        content = item.find('div', class_='record-body').text.strip()
        news_list.append((link, title, content))
    # for news in news_list:
    #     print(news)

    # print("-"*30)
    # if len(news_list) > 0:
    #     print(f'Pierwsza wiadmość: \nLink: {news_list[0][0]}\nTytuł: {news_list[0][1]}\nTreść: {news_list[0][2]}')
    # else: 
    #     print("Brak wiadomości dla wybranego podmiotu")
    # print("-"*30)
    return news_list

# prace w trakcie ...
def last_report(ticker):
    pass

# zwraca listę dat i listę przychodów od najstarszego do najnowszego
def chart_data(ticker):
    daty = []
    przychody = []  
    tir = str(ticker)
    if tir[-1] == 'S' and tir[-2] == 'U' and tir[-3] == '.':
        print("TAK! BINOGO, AMERYKA!")
    strona = "https://www.biznesradar.pl/raporty-finansowe-rachunek-zyskow-i-strat/" + ticker
    driver = driver_settings()   
    try:
        driver.get(strona)
        print("Udało sie połączyć")
        time.sleep(2)
    except:
        print(f"Nie udało się otworzyć strony - {ticker}")
        driver.quit()
    
    try:
        element = driver.find_element(By.CLASS_NAME, 'report-table')
        content = element.get_attribute('innerHTML')
        teksthtml = ''.join(content)
        soup = BeautifulSoup(teksthtml, 'html.parser')
        rok = soup.find_all('th', class_='thq h')
        rok_najnowszy = soup.find_all('th', class_='thq h newest')
        elementA = [item.text.split('\n')[1] for item in rok_najnowszy]
        elements_A = [item.text.split('\n')[1] for item in rok]
        for i in elements_A:
            daty.append(int(i))
        try:
            last = int(elementA[-1])
        except: 
            last = int(elements_A[-1])+1
        daty.append(last)
        zysk = soup.find_all('span', class_='pv')
        _ = 0
        for i in zysk:
            if not '%' in i.text and _ < len(daty):
                _+=1
                przychody.append(i.text)
        if len(przychody) == len(daty):
            print("Dane są kompletne")
        else:
            raise Exception("Dane są niekompletne")
        driver.quit()
        return daty, przychody

    except Exception as e:
        print(f'Błąd... nie znaleziono szukanego elementu:\nException - {e}')
        driver.quit()

# Główna funkcja odpalająca funkcjonalności
def main():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    creds = Credentials.from_service_account_file(path, scopes=scopes)
    client = gspread.authorize(creds)

    sheetID = "1CKlpkb9neJhdKAfpsK7Jxizue9XcEm5kquSkLjqh1Yc"
    workbook = client.open_by_key(sheetID)
    sheet = workbook.worksheet("Dashboard")
    valuse = sheet.get('A1')
    ticker = valuse[0][0]
    wskazniki = scrap_financial_data(ticker)
    news = scrap_news(ticker)
    daty, przychody = chart_data(ticker)
    for i in range(len(wskazniki)):
        sheet.update_cell(i+100, 1, wskazniki[i])
    for i in range(3):
        sheet.update_cell(i+100, 4, news[0][i])

    row_numbers = [139, 140]
    for row_number in row_numbers:
        row_cells = sheet.row_values(row_number)  # Pobierz wszystkie wartości w wierszu
        last_col_letter = column_number_to_letter(len(row_cells))  # Ostatnia kolumna w formacie litery Excel
        range_notation = f"A{row_number}:{last_col_letter}{row_number}"  # Określ poprawny zakres, np. "A15:D15"
        empty_values = [''] * len(row_cells)  # Tworzy listę pustych wartości o długości wiersza
        sheet.update([empty_values], range_notation)  # Aktualizuje wiersz na puste wartości

    for i in range(len(daty)):
        sheet.update_cell(139, 1+i, daty[i])
        sheet.update_cell(140, 1+i, przychody[i])


    print('zaktualizowano')


main()
