"""
Taking top 250 films from IMDB site and using bs4 to parse the page then save 
those results in an excel workbook.

"""
from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
#Name of the sheet
sheet.title = 'Movies'
sheet.append(['Movie Name','year','Rating'])
try:
    response = requests.get('https://www.imdb.com/chart/top/')
    soup = BeautifulSoup(response.text, 'html.parser')
    movies = soup.find('tbody',class_='lister-list').find_all('tr')
    for movie in movies:
        movie_name = movie.find('td',class_ = 'titleColumn').a.text
        year_movie = movie.find('td',class_ = 'titleColumn').span.text
        movie_rating = movie.find('td', class_ = 'ratingColumn imdbRating').strong.text
        sheet.append([movie_name,year_movie,movie_rating])
except Exception as e:
    print(e)

excel.save('C:\\Users\\Filename.xlsx')

