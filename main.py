from bs4 import BeautifulSoup
import requests
from openpyxl import load_workbook
import os


def parse1():
    filename = 'лабалоторная1.xlsx'
    exelfile = load_workbook(filename)
    spisok = exelfile['данные']
    description1 = []
    description2 = []
    description3 = []
    count = 2
    ssilka = 'https://www.chitai-gorod.ru/catalog/collections/bestsell?page='

    for i in range(1, count + 1):
        url1 = ssilka + str(i)
        page1 = requests.get(url1)
        print(page1.status_code)
        soup1 = BeautifulSoup(page1.text, "html.parser")
        block1 = soup1.findAll('article', class_='product-card product-card product')

        for data in block1:
            name = data.find(class_='product-title__head')
            author = data.find(class_='product-title__author')
            price = data.find(class_='product-price__value')
            if (name and author and price) is not None:
                description1.append(name.text)
                description2.append(author.text)
                description3.append(price.text)

    print(description1, description2, description3)
    pozitionNow = 1
    for elem1, elem2, elem3 in zip(description1, description2, description3):
        cell = spisok.cell(1, pozitionNow)
        cell.value = elem1

        cell = spisok.cell(2, pozitionNow)
        cell.value = elem2

        cell = spisok.cell(3, pozitionNow)
        cell.value = elem3

        pozitionNow += 1

    exelfile.save(filename)
    exelfile        .close()

    # Открывает файл Excel после сохранения
    os.startfile(filename)


parse1()
