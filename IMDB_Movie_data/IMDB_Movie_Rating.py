# Importing necessary libraries
# pip install bs4 requests openpyxl html5lib
import requests, openpyxl
from bs4 import BeautifulSoup

# Creating excel workbook
file = openpyxl.Workbook()
sheet = file.active
sheet.title = "Top Rated Movies"
sheet.append(["Rank", "Movie Name", "Year of Release", "IMDB Rating"])

# Accessing the website
url = "https://www.imdb.com/chart/top/"
site = requests.get(url)
soup = BeautifulSoup(site.text, "html5lib")

for movie in soup.find('tbody', class_="lister-list").find_all('tr'):
    rank = int(movie.find('td', class_="titleColumn").get_text(strip=True).split(".")[0])
    name = movie.find('td', class_="titleColumn").a.text
    year = int(movie.find('td', class_="titleColumn").span.text.strip("()"))
    rating = float(movie.find('td', class_="ratingColumn imdbRating").strong.text)

    sheet.append([rank, name, year, rating]) 

# Saving the excel file
file.save("IMDB Ratings data.xlsx")   