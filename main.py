import re
import timeit
from datetime import date
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def calculate_runtime(func):
    def inner1(*args, **kwargs):
        begin = timeit.default_timer()
        return_value = func(*args, **kwargs)
        end = timeit.default_timer()
        print("Runtime of '" + func.__name__ + "' function : " + str(end - begin))
        return return_value
    return inner1


@calculate_runtime
def scrape(ws, today): 

    col_idx = get_column_letter(ws.max_column + 2)
    html_text = request_appropriate_website(ws)

    if html_text == 'error':
        add_formula(ws, col_idx)
        # Return statement
        return "Skipped"

    soup = BeautifulSoup(html_text, 'lxml')
    ranking = soup.find_all('tr', class_ = 'ranking-list')

    for i in ranking:
        temp = get_appropriate_data(ws, i)
        add_to_current_worksheet(ws, temp[0], float(temp[1]), col_idx)

    ws.auto_filter.ref = "A1:" + col_idx + str(ws.max_row)
    ws[col_idx + '1'] = today
    ws[chr(ord(col_idx) - 1) + '1'] = 'Change'

    return "Executed"

#TODO replace with dictionary
def request_appropriate_website(ws):
    temp = get_data(ws.title)
    
    if temp == 'error':
        return temp

    return requests.get(temp[1]).text


def get_data(x):
    TITLE_DICTIONARY = {
        'ARV': ('ARO', 'https://myanimelist.net/topanime.php'),
        'AMV': ('AMO', 'https://myanimelist.net/topanime.php?type=bypopularity'),
        'AFV': ('AFO', 'https://myanimelist.net/topanime.php?type=favorite'),
        'MRV': ('MRO', 'https://myanimelist.net/topmanga.php'),
        'MMV': ('MMO', 'https://myanimelist.net/topmanga.php?type=bypopularity'),
        'MFV': ('MFO', 'https://myanimelist.net/topmanga.php?type=favorite'),
    }
    return TITLE_DICTIONARY.get(x, 'error')


#TODO replace with dictionary
def get_appropriate_data(ws, data):
    if ws.title == 'ARV':
        title = data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        core = data.find('td', class_ = 'score ac fs14').text.replace('\n','').replace(' ','')
    if ws.title == 'AMV':
        title = data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        temp = data.find(string = re.compile('members'))
        core = temp.replace(' ','').replace(',','').replace('members','').replace('\n','')
    if ws.title == 'AFV':
        title = data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        temp = data.find(string = re.compile('favorites'))
        core = temp.replace(' ','').replace(',','').replace('favorites','').replace('\n','')
    if ws.title == 'MRV':
        title = data.find('h3', class_ = 'manga_h3').text
        core = data.find('td', class_ = 'score ac fs14').text.replace('\n','')
    if ws.title == 'MMV':
        title = data.find('h3', class_ = 'manga_h3').text
        temp = data.find(string = re.compile('members'))
        core = temp.replace(' ','').replace(',','').replace('members','').replace('\n','')
    if ws.title == 'MFV':
        title = data.find('h3', class_ = 'manga_h3').text
        temp = data.find(string = re.compile('favorites'))
        core = temp.replace(' ','').replace(',','').replace('favorites','').replace('\n','')
    return title, core


def add_to_current_worksheet(ws, title, info, col_idx):
    rows = row_count(ws)
    for i in range(2, rows):
        if ws['B' + str(i)].value == title:
            ws[col_idx + str(i)] = info
            if ws[chr(ord(col_idx) - 2) + str(i)].value:
                ws[chr(ord(col_idx) - 1) + str(i)].value = info - ws[chr(ord(col_idx) - 2) + str(i)].value
            break
    else:
        ws['A' + str(rows)] = '#' + str(rows - 1)
        ws['B' + str(rows)] = title
        ws[col_idx + str(rows)] = info


def row_count(ws):
    row_count = 1
    while ws['A' + str(row_count)].value:
            row_count += 1
    return row_count


def add_formula(ws, col_idx):
    for i in range(2, 52):
        ws[col_idx + str(i)].value = '=COUNTIF(ARV!$' + col_idx + '$' + str(i) + ':$' + col_idx + '$100,">"&ARV!' + col_idx + str(i) + ')+1'


def main():
    wb = load_workbook("./Excel/Input.xlsx")
    temp = date.today()
    TODAYS_DATE = temp.strftime("%Y.%m.%d")

    for ws in wb:
        print(scrape(ws, TODAYS_DATE))

    start = timeit.default_timer()
    wb.save("./Excel/" + TODAYS_DATE + ".xlsx")
    end = timeit.default_timer()
    print("Saved in " + str(end - start))


if __name__ == '__main__':
    main()