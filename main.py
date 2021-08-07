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
        print("Runtime of '" + func.__name__ + "' function : " + str(end - begin) + "\n")
        return return_value
    return inner1


def get_worksheet_title(worksheet_title):
    WORKSHEET_TITLE_DICTIONARY = {
        'ARV': 'ARO',
        'AMV': 'AMO',
        'AFV': 'AFO',
        'MRV': 'MRO',
        'MMV': 'MMO',
        'MFV': 'MFO'
    }
    return WORKSHEET_TITLE_DICTIONARY.get(worksheet_title)


def get_url(worksheet_title):
    URL_DICTIONARY = {
        'ARV': 'https://myanimelist.net/topanime.php',
        'AMV': 'https://myanimelist.net/topanime.php?type=bypopularity',
        'AFV': 'https://myanimelist.net/topanime.php?type=favorite',
        'MRV': 'https://myanimelist.net/topmanga.php',
        'MMV': 'https://myanimelist.net/topmanga.php?type=bypopularity',
        'MFV': 'https://myanimelist.net/topmanga.php?type=favorite'
    }
    return URL_DICTIONARY.get(worksheet_title)


def request_appropriate_website(ws):
    return requests.get(get_url(ws.title)).text


def get_appropriate_data(ws, soup_data):
    if ws.title == 'ARV':
        title = soup_data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        core = soup_data.find('td', class_ = 'score ac fs14').text.replace('\n','').replace(' ','')
    if ws.title == 'AMV':
        title = soup_data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        temp = soup_data.find(string = re.compile('members'))
        core = temp.replace(' ','').replace(',','').replace('members','').replace('\n','')
    if ws.title == 'AFV':
        title = soup_data.find('h3', class_ = 'hoverinfo_trigger fl-l fs14 fw-b anime_ranking_h3').text
        temp = soup_data.find(string = re.compile('favorites'))
        core = temp.replace(' ','').replace(',','').replace('favorites','').replace('\n','')
    if ws.title == 'MRV':
        title = soup_data.find('h3', class_ = 'manga_h3').text
        core = soup_data.find('td', class_ = 'score ac fs14').text.replace('\n','')
    if ws.title == 'MMV':
        title = soup_data.find('h3', class_ = 'manga_h3').text
        temp = soup_data.find(string = re.compile('members'))
        core = temp.replace(' ','').replace(',','').replace('members','').replace('\n','')
    if ws.title == 'MFV':
        title = soup_data.find('h3', class_ = 'manga_h3').text
        temp = soup_data.find(string = re.compile('favorites'))
        core = temp.replace(' ','').replace(',','').replace('favorites','').replace('\n','')
    return (title, core)


@calculate_runtime
def scrape(value_ws, order_ws): 

    collumn_index = get_column_letter(value_ws.max_column + 2)
    html_text = request_appropriate_website(value_ws)

    soup = BeautifulSoup(html_text, 'lxml')
    ranking_list = soup.find_all('tr', class_ = 'ranking-list')

    for animanga in ranking_list:
        temp = get_appropriate_data(value_ws, animanga)
        add_data_to_worksheets(value_ws, order_ws, temp[0], float(temp[1]), collumn_index)

    value_ws.auto_filter.ref = "A1:" + collumn_index + str(value_ws.max_row)
    value_ws[collumn_index + '1'] = date.today()
    value_ws[chr(ord(collumn_index) - 1) + '1'] = 'Change'

    order_ws.auto_filter.ref = "A1:" + collumn_index + str(value_ws.max_row)
    order_ws[collumn_index + '1'] = date.today()
    order_ws[chr(ord(collumn_index) - 1) + '1'] = 'Change'


def get_worksheet_row_count(ws):
    row_count = 1
    while ws['A' + str(row_count)].value:
            row_count += 1
    return row_count


def add_data_to_worksheets(v_ws, o_ws, title, info, col_idx):
    for i in range(2, get_worksheet_row_count(v_ws)):
        if v_ws['B' + str(i)].value == title:
            v_ws[col_idx + str(i)] = info
            if v_ws[chr(ord(col_idx) - 2) + str(i)].value:
                v_ws[chr(ord(col_idx) - 1) + str(i)].value = info - v_ws[chr(ord(col_idx) - 2) + str(i)].value

            o_ws[col_idx + str(i)] = '=COUNTIF(' + v_ws.title + '!$' + col_idx + '$2:$' + col_idx + '$100,">"&' + v_ws.title + '!' + col_idx + str(i) + ')+1'
            if o_ws[chr(ord(col_idx) - 2) + str(i)].value:
                o_ws[chr(ord(col_idx) - 1) + str(i)].value = '=' + (chr(ord(col_idx) - 2) + str(i)) + ' - ' + (col_idx + str(i))

            break

    else:
        print('New animanga added to ' + v_ws.title + ': ' + title)
        ws_row_count = get_worksheet_row_count(v_ws)

        v_ws['A' + str(ws_row_count)] = '#' + str(ws_row_count - 1)
        v_ws['B' + str(ws_row_count)] = title
        v_ws[col_idx + str(ws_row_count)] = info

        o_ws['A' + str(ws_row_count)] = '#' + str(ws_row_count - 1)
        o_ws['B' + str(ws_row_count)] = title
        o_ws[col_idx + str(ws_row_count)] = '=COUNTIF(' + v_ws.title + '!$' + col_idx + '$2:$' + col_idx + '$100,">"&' + v_ws.title + '!' + col_idx + str(ws_row_count) + ')+1'


def main():
    workbook = load_workbook("./Excel/Input.xlsx")

    WORKSHEET_TITLE_DICTIONARY = {
        'ARV': 'ARO',
        'AMV': 'AMO',
        'AFV': 'AFO',
        'MRV': 'MRO',
        'MMV': 'MMO',
        'MFV': 'MFO'
    }
    
    temp = date.today()
    TODAYS_DATE = temp.strftime("%Y.%m.%d")

    for main_worksheet_title in WORKSHEET_TITLE_DICTIONARY:
        scrape(workbook[main_worksheet_title], workbook[WORKSHEET_TITLE_DICTIONARY[main_worksheet_title]])

    start = timeit.default_timer()
    workbook.save("./Excel/" + TODAYS_DATE + ".xlsx")
    end = timeit.default_timer()
    print("Saved as '" + TODAYS_DATE + ".xlsx" + "' in " + str(end - start))


if __name__ == '__main__':
    main()