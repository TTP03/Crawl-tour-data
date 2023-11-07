import os
from bs4 import BeautifulSoup
from openpyxl import Workbook
City = []
Category = []
Title = []
Price = []
Rating = []
Time = []
Img_Src = []
Detail = []

cityPath = "City/"
cityList = os.listdir(cityPath)

for city in cityList:

    categoryPath = cityPath + city +'/'
    categoryList = os.listdir(categoryPath)

    for category in categoryList:

        if '.html' in category:

            with open(categoryPath + category , 'r', encoding='utf-8') as file:
                html_content = file.read()

                soup = BeautifulSoup(html_content, 'html.parser')
                block = soup.find_all('div', class_='activity-card-block activity-card-block--grid')
                for i in block:
                    City.append(city)

                    Category.append(category[:-5])

                    title = i.find('p', class_='vertical-activity-card__title').text.strip()
                    Title.append(title)

                    price = i.find('span', class_='baseline-pricing__from--value')
                    if price:
                        Price.append(price.text.strip())
                    else:
                        Price.append(None)

                    rating = i.find('span', class_='rating-overall__rating-number rating-overall__rating-number--right')
                    if rating: 
                        Rating.append(rating.text.strip())
                    else:
                        Rating.append(None)

                    time = i.find('span', class_='bullet')
                    if time: 
                        Time.append(time.text.strip())
                    else:
                        Time.append(None)

                    images = i.find('img')
                    if images.get('data-src'):
                        img = images.get('data-src')
                    else:
                        img = images.get('src')
                    Img_Src.append(img)

                    link = i.find('a')
                    Detail.append(link.get('href'))

workbook = Workbook()
sheet = workbook.active
sheet.append(['City', 'Category', 'Title', 'Price', 'Rating', 'Time', 'Image source', 'Detail'])
for a,b,c,d,e,f,g,h in zip(City, Category, Title, Price, Rating, Time, Img_Src, Detail):
    sheet.append([a,b,c,d,e,f,g,h])
workbook.save('Tour Data.xlsx')