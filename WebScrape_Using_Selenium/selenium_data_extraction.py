# WebDriver download link --> https://sites.google.com/a/chromium.org/chromedriver/downloads
# pip install selenium tkinter openpyxl
# Importing necessary libraries
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from tkinter import * 
from tkinter import messagebox as msg
import openpyxl

# Creating a workbook
book = openpyxl.Workbook()
sheet = book.active
sheet.title = "Top rated movies"
sheet.append(["Rank", "Movie Name", "Year of Release", "IMDB Ratings"])

# initialising the chrome webdriver and automating the websites
driver = webdriver.Chrome("D:/Softwares/chromedriver.exe")
driver.maximize_window()
driver.get("https://www.google.com/")

search = driver.find_element(By.NAME, "q")
search.send_keys("imdb top movies")
search.send_keys(Keys.RETURN)

link = driver.find_element(By.XPATH, "//a[@href='https://www.imdb.com/chart/top/']")
link.send_keys(Keys.RETURN)

try:
    tbody = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "tbody")))
    movies = tbody.find_elements(By.TAG_NAME, "tr")

    # Getting the required data
    for Yscroll, movie in zip(range(0, 19000, 70), movies):
        driver.execute_script(f"window.scrollTo(0, {Yscroll})")
        movieInfo = movie.find_element(By.CLASS_NAME, "titleColumn")

        rank = movieInfo.text.split(".")[0]
        name = movieInfo.find_element(By.TAG_NAME, "a").text
        year = movieInfo.find_element(By.CLASS_NAME, "secondaryInfo").text.strip("()")
        ratings = movie.find_element(By.TAG_NAME, "strong").text

        sheet.append([int(rank), name, int(year), float(ratings)]) 
        
    # Saving the excel file 
    book.save("Top rated movies data.xlsx")
    
finally:
    driver.quit()
    root = Tk()
    root.wm_withdraw()
    msg.showinfo("INFO", "Data Extracted Successfully!!...")
