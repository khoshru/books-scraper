import time
import requests
from bs4 import BeautifulSoup
import openpyxl

BASE_URL = "https://books.toscrape.com/catalogue/page-{}.html"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
}

results = []

for page in range(1, 6):
    url = BASE_URL.format(page)

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
    except Exception as e:
        print(f"مشکل در صفحه {page}: {e}")
        continue

    if response.status_code != 200:
        print(f"خطای {response.status_code} در صفحه {page}")
        continue

    response.encoding = "utf-8"
    soup = BeautifulSoup(response.text, "html.parser")
    books = soup.find_all("article", class_="product_pod")

    if not books:
        print("کتابی پیدا نشد — احتمالاً آخرین صفحه")
        break

    print(f"page: {page} book: {len(books)}")

    for book in books:
        try:
            title = book.find("h3").find("a")["title"]
            price = book.find("p", class_="price_color").text
            image = book.find("img")["src"]
        except AttributeError:
            continue

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