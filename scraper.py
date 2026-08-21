import time
import requests
from bs4 import BeautifulSoup
import openpyxl

results = []

for page in range(1, 6):
    url = f"https://books.toscrape.com/catalogue/page-{page}.html"
    response = requests.get(url)
    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")
    books = soup.find_all("article", class_="product_pod")

    print(f"page: {page} book: {len(books)}")

    for book in books:
        title = book.find("h3").find("a")["title"]
        price = book.find("p", class_="price_color").text
        image = book.find("img")["src"]

        results.append({
            "title": title,
            "price": price,
            "image": image
        })

    time.sleep(1)

print(f"total: {len(results)}")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Books"

ws.append(["Title", "Price", "Image"])

for b in results:
    ws.append([b["title"], b["price"], b["image"]])

wb.save("books.xlsx")
print("done")




