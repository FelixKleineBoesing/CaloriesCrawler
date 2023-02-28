# This program scrapes the information from fddb.info to get calories and other informations about food items
import time

import xlwings as xw
import random
from foodcrawler.parsing import get_food_informations


def main():
    workbook_name = "Rezepte.xlsm"
    wb: xw.Book = xw.books[workbook_name]
    ws: xw.Sheet = wb.sheets["ZutatenListe"]
    num_row = ws.range('A1').end('down').row

    # collect data
    content_list = ws.range((2, 1), (num_row, 2)).value
    print(content_list)
    food_items = {}
    for i, content in enumerate(content_list):
        original_name, fddb_name = content
        if fddb_name is None:
            informations = get_food_informations(content[0])
            food_items[i + 2] = informations
            time.sleep(random.randint(1, 10) / 10)
    for row, infos in food_items.items():
        ws.range((row, 2)).value = infos["product_name"]
        ws.range((row, 3)).value = infos["serving_value"]
        ws.range((row, 4)).value = infos["serving_unit"]
        ws.range((row, 5)).value = infos["calories"]
        ws.range((row, 6)).value = infos["protein"]
        ws.range((row, 7)).value = infos["fat"]
        ws.range((row, 8)).value = infos["carbs"]
        ws.range((row, 9)).value = infos["relation"]
        ws.range((row, 10)).value = infos["product_link"]

if __name__ == "__main__":
    main()

