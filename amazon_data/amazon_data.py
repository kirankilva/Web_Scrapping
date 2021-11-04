# Importing required libraries
# pip install BeautifulSoup
# pip install requests
# pip install pandas
from bs4 import BeautifulSoup 
import requests                   
import pandas as pd               


reviewlst = []
# Creating function for extracting data from given link
def get_review_info():
    for tag in soup.find_all('div', attrs={'data-hook':'review'}):
        try:
            review = {
                    'reviewer':tag.find('span', {'class':'a-profile-name'}).string,
                    'posted_date':tag.find('span', {'data-hook':'review-date'}).text.replace('Reviewed in India on ', '').strip(),
                    'review_text':tag.find('span', {'data-hook':'review-body'}).text.strip()
                    }
            reviewlst.append(review)
        except:
            pass

# Looping through the pages
for i in range(1,50):
    url = f"https://www.amazon.in/LG-24-inch-Monitor-Freesync-Borderless/product-reviews/B08J5Y9ZSV/ref=cm_cr_arp_d_paging_btm_next_2?ie=UTF8&reviewerType=all_reviews&pageNumber={i}"
    site = requests.get(url)
    soup = BeautifulSoup(site.content, 'html5lib')
    get_review_info()     # Function call


# Creating DataFrame using pandas
df = pd.DataFrame(reviewlst)

# Exporting DataFrame to Excel sheet
df.to_excel("amazon_data.xlsx", index=False)
