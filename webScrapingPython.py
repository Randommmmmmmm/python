from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active


print('Web Scraping in IMDB and exporting results to Excel\n')
category = {1:'Top 250 Movies of all time',2:'Top Action Movies'}
for key,value in category.items():
    print(key, '-' ,value)
    
choice = int(input('Select  your choice: '))

if choice in category.keys():

     if choice == 1:
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

        excel.save('C:\\Users\\home\\Desktop\\Filename.xlsx')
        print('File generated successfully!!')

     if choice == 2:
        sheet.title = 'Action movie list'
        sheet.append(['index','Movie Name','Rating','Description','Director Name','Gross Volume'])
        try:
            response = requests.get('https://www.imdb.com/search/title/?genres=action&sort=user_rating,desc&title_type=feature&num_votes=25000,')
            soup = BeautifulSoup(response.text, 'html.parser')
            movies = soup.find('div',class_='lister-list').find_all('div',class_ = 'lister-item')
            for movie in movies:
                index = movie.find('h3').find('span',class_ = 'lister-item-index').text.split('.')[0]
                movie_name = movie.find('h3').a.text
                rating = movie.find('div',class_ = 'ratings-bar').find('strong').text
                description = movie.find('p', class_ = 'text-muted').findNext('p').text
                director_name = movie.find('p',class_ = '').a.text
                gross = movie.find('p', class_ = 'sort-num_votes-visible').find_all('span')[-1].text
                sheet.append([index,movie_name,rating, description,director_name,gross])
           
        except Exception as e:
            print(e)

        try:
            excel.save('C:\\Users\\home\\Action.xlsx')
            print('File generated successfully!!')
        except Exception as ee:
            print(ee)
        

else:
    print('Please select only from the above given list')


    
