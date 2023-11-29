from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rared Movies'
sheet.append(['Rank', 'Name', 'Year', 'IMDb'])

url = 'https://www.imdb.com/chart/top/?ref_=nv_mv_250'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}
response = requests.get(url, headers=headers)
try:
    source = response
    source.raise_for_status()
    soup = BeautifulSoup(source.content, 'html.parser')
    movies = soup.find_all('div', class_="sc-479faa3c-0 fMoWnh cli-children")

    # Writing the list of movies in a HTML file to be easier to scrape
    with open("scraping-imdb/data/imdb-list.html", "w", encoding='utf8') as file:
        for movie in movies:
            file.write(movie.prettify())

    # Separating the data for rank, name, year and rating
    for movie in movies:
        rank, name = str(movie.find('h3', class_="ipc-title__text").text).split(". ")
        year = movie.find('span', class_="sc-479faa3c-8 bNrEFi cli-title-metadata-item").text
        rating = str(movie.find('span', class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text).split('(')[0].strip()
        
        sheet.append([rank, name, year, rating])

    # Saving the values in a xlsx file
    excel.save('scraping-imdb/data/imdb-movies-rating.xlsx')
except Exception as e:
    print(e)
