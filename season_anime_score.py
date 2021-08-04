from bs4 import BeautifulSoup
import requests

html_text = requests.get('https://myanimelist.net/anime/season').text
soup = BeautifulSoup(html_text, 'lxml')

animes = soup.find_all('div', class_ = 'seasonal-anime js-seasonal-anime')

for anime in animes:
    title = anime.find('h2', class_ = 'h2_anime_title').text
    score = anime.find('span', title = 'Score').text.replace(' ','').replace('\n','')
    members = anime.find('span', class_ = 'member fl-r').text.replace(',','').replace(' ','').replace('\n','')

    if int(members) < 150000:
        break

    print(title)
    print(score)
    print(members)

    print('----------------')