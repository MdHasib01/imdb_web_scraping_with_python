#-----------------------------------------------------------
#       1. Web scraping with python.
#       2. Saving the data to a excel file.
#       3. two package must needed: BeautifulSoup, openpyxl
#       4. install requests,bs4,openpyxl
#-----------------------------------------------------------

from bs4 import BeautifulSoup;
import requests, openpyxl
# Creating excel file with OpenPyXl
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top imdb movie list"
sheet.append(["Rank","Movie Name","Year","Ratings"])

# Getting data form the website
try:
    source = requests.get("https://www.imdb.com/chart/top")
    source.raise_for_status()

    soup = BeautifulSoup(source.text, "html.parser")
    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    
    # Looping the movie list
    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip("()")
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print([rank,name,year,rating])
        # Adding the movie data to the excel file
        sheet.append([rank,name,year,rating])
        # break

except Exception as e:
    print(e)
# saving the excel file as xlsx format
excel.save('IMDB movie rating.xlsx')