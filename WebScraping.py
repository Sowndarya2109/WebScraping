from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Name', 'Year of Release', 'Movie Rating'])

try:
    URL = 'https://m.imdb.com/chart/top/'
    HEADER = {
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.8 Safari/537.36",
        'Accept-Language': "en-US,en;q=0.5"
    }

    source = requests.get(URL, headers=HEADER)
    source.raise_for_status() 
    soup = BeautifulSoup(source.text, 'html.parser')

    movies = soup.find('ul', class_='ipc-metadata-list').find_all('li')

    for movie in movies:
        name = movie.find('span', class_='ipc-title__text').text.strip()
        year = movie.find('span', class_='sc-8c396aa2-2 itZqyK').text.strip('()')
        rating = movie.find('span', class_='ipc-rating-star').text.strip()

        print(name, year, rating)

        sheet.append([name, year, rating])

except Exception as e:
    print(f"Error: {e}")

excel.save('Movies.xlsx')
