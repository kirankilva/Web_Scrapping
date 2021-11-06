# pip install bs4 requests openpyxl
# Importing necessary libraries
import requests, openpyxl
from bs4 import BeautifulSoup


# Creating a excel workbook
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "Top 1000 Movies data"
sheet.append(["Rank", "Movie Name", "Year of Released", "Certificate", "Genre", "IMDB Ratings", "Movie Director", "Gross Amount"])

try:
    # Accessing the website
    url = "https://www.imdb.com/search/title/?groups=top_1000&sort=user_rating,desc&count=100&ref_=adv_prv"
    site = requests.get(url)
    site.raise_for_status()    # Raises an error when url is invalid in except block
    soup = BeautifulSoup(site.text, "html.parser")

    certificate_text = ""
    gross_text = ""
    # Defining a function
    def movies_data():
        global text_data
        global gross_data
        for movie in soup.find('div', class_="lister-list").find_all('div', class_="lister-item mode-advanced"):
            tag = movie.find('div', class_="lister-item-content")

            rank = tag.find('h3', class_="lister-item-header").span.text.replace(".","").replace(",", "")
            name = tag.find('h3', class_="lister-item-header").a.text
            year = tag.find('h3', class_="lister-item-header").find('span', class_="lister-item-year text-muted unbold").text.split(" ")[-1].strip("()")

            certificate = tag.find('p', class_="text-muted").find('span', class_="certificate")
            if certificate == None:
                certificate_text = "Unknown"
            else:
                certificate_text = certificate.text

            genre = tag.find('p', class_="text-muted").find('span', class_="genre").text.strip()
            rating = tag.find('div', class_="ratings-bar").find('div', class_="inline-block ratings-imdb-rating").strong.text
            director = tag.find('p', class_="").a.text

            gross = tag.find('p', class_="sort-num_votes-visible").text.strip().split("\n")[-1]
            if gross[0] == "$":
                gross_text = gross
            else:
                gross_text = "Unknown"
            
            # Adding data to the excel workbook
            sheet.append([int(rank), name, int(year), certificate_text, genre, float(rating), director, gross_text])
            

    # Calling a function for first 100 Movies only
    movies_data()

    # Calling a function through loop except first 100 Movies
    for i in range(101, 1001, 100):
        url = f"https://www.imdb.com/search/title/?groups=top_1000&sort=user_rating,desc&count=100&start={i}&ref_=adv_nxt"
        site = requests.get(url)
        site.raise_for_status()
        soup = BeautifulSoup(site.text, "html.parser")
        movies_data()
        print(f"Finished page {i}")

    # Saving the excel Workbook
    book.save("Top 1000 IMDB Rating Movies data.xlsx")


except Exception as e:
    print(e)
