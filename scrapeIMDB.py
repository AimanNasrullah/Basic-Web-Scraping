from bs4 import BeautifulSoup
import requests, openpyxl

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Rank', 'Movie Title', 'Year of Release', 'IMDB Rating'])

try:
  headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
  source = requests.get('https://www.imdb.com/chart/top/?ref_=nv_mv_250', headers=headers)

  soup = BeautifulSoup(source.text, 'lxml')

  movies = soup.find('ul', class_='ipc-metadata-list ipc-metadata-list--dividers-between sc-3f13560f-0 sTTRj compact-list-view ipc-metadata-list--base').find_all('li')
  for movie in movies:
    movie_title = movie.find('h3', class_='ipc-title__text').text
    name = movie_title.split('.')[1]
    rank = movie_title.split('.')[0]
    year = movie.find('span', class_='sc-b85248f1-6 bnDqKN cli-title-metadata-item').text.strip('()')
    rating_span = movie.find('span', class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').get_text(strip=True)
    rating = rating_span.split('(')[0]
    formatted_title = f'{movie_title} ({year}) {rating}'
    print(rank, name, year, rating)
    sheet.append([rank, name, year, rating])
  #for index, movie in enumerate(movies):
    #fullname = movie.find('h3', class_='ipc-title__text').text
    #fullname = fullname.split(".") #'Eg. 1. The Shawshank Redemption',but we want 'The Shawshank Redemption' only
      # Extract the movie title only
    #movie_title = fullname[1].strip()
    #print(movie_title)

except Exception as e:
  print(e)

excel.save('IMDB Top Rated Movies.csv')