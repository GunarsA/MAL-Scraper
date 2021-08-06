import re
import timeit
from datetime import date
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def scrape():
    # Start timer
    start = timeit.default_timer()

    # Get first empty column 
    col_idx = get_column_letter(ws.max_column + 2)

    # Get html using requests
    html_text = request_website()
    # Check if correct spreadsheet
    if not html_text:
        add_formula(col_idx)
        # Return statement
        return "Skipped"
    # Parse html using lxml
    soup = BeautifulSoup(html_text, 'lxml')
    # Get data
    ranking = soup.find_all('tr', class_ = 'ranking-list')

    # Loop through entries
    for i in ranking:
        # Save data in temperary tuple
        temp = find_core(i)
        # Add data to worksheet
        add_to_current_ws(temp[0], float(temp[1]), col_idx)

    # Add auto filter
    ws.auto_filter.ref = "A1:" + col_idx + str(ws.max_row)

    # Add date rows title
    ws[col_idx + '1'] = today
    # Add 'Change' rows title
    ws[chr(ord(col_idx) - 1) + '1'] = 'Change'
    # End timer
    end = timeit.default_timer()

    # Return execution time
    return "Executed in " + str(end - start)

def request_website():
    # Get specific website's adress
    if ws.title == 'ARV':
        return requests.get('https://myanimelist.net/topanime.php').text
    if ws.title == 'AMV':
        return requests.get('https://myanimelist.net/topanime.php?type=bypopularity').text
    if ws.title == 'AFV':
        return requests.get('https://myanimelist.net/topanime.php?type=favorite').text
    if ws.title == 'MRV':
        return requests.get('https://myanimelist.net/topmanga.php').text
    if ws.title == 'MMV':
        return requests.get('https://myanimelist.net/topmanga.php?type=bypopularity').text
    if ws.title == 'MFV':
        return requests.get('https://myanimelist.net/topmanga.php?type=favorite').text


def find_core(data):
    # Get specific website's data
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
    # Return as tuple
    return title, core

def add_to_current_ws(title, info, col_idx):
    # Get row count
    rows = row_count()
    # Parse all entries
    for i in range(2, rows):
        # Find if entry already exists
        if ws['B' + str(i)].value == title:
            # Add new data
            ws[col_idx + str(i)] = info
            # Check if there is data in previous column
            if ws[chr(ord(col_idx) - 2) + str(i)].value:
                # Calculate change from previous entry
                ws[chr(ord(col_idx) - 1) + str(i)].value = info - ws[chr(ord(col_idx) - 2) + str(i)].value
            break
    # If no matching entry found add new
    else:
        # Add index
        ws['A' + str(rows)] = '#' + str(rows - 1)
        # Add title
        ws['B' + str(rows)] = title
        # Add data
        ws[col_idx + str(rows)] = info

def row_count():
    # Declare integer
    row_count = 1
    # Count rows with data
    while ws['A' + str(row_count)].value:
            row_count += 1
    # Return row count + 1
    return row_count

def add_formula(col_idx):
    for i in range(2, 52):
        ws[col_idx + str(i)].value = '=COUNTIF(ARV!$' + col_idx + '$' + str(i) + ':$' + col_idx + '$100,">"&ARV!' + col_idx + str(i) + ')+1'

def main():
    # Load workbook
    wb = load_workbook("./Excel/Input.xlsx")

    # Get todays date
    today = date.today()
    # Convert date to #YY.MM.DD format
    date = today.strftime("%Y.%m.%d")

    # Loop through worksheets
    for ws in wb:
        # Add new data to spreadsheet
        print(scrape())

    # Start timer
    start = timeit.default_timer()
    # Save workbook
    wb.save("./Excel/" + date + ".xlsx")
    # End timer
    end = timeit.default_timer()
    # Output save time
    print("Saved in " + str(end - start))

if __name__ == '__main__':
    main()