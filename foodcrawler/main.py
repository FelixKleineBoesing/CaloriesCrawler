# This program scrapes the information from fddb.info to get calories and other informations about food items
import logging
import time
import sys
import xlwings as xw
import random
from tqdm import tqdm
from foodcrawler.parsing import get_food_informations


def main():
    assert len(sys.argv) >= 2, "Please provide the workbook name as argument top the executable"
    workbook_name = sys.argv[1]
    print("Starting the crawling for workbook: {}".format(workbook_name))
    wb: xw.Book = xw.books[workbook_name]
    ws: xw.Sheet = wb.sheets["Ingredients"]
    num_row = ws.range('A1').end('down').row
    print("Found {} rows in the sheet".format(num_row))
    # collect data
    content_list = ws.range((2, 1), (num_row, 2)).value
    food_items = {}
    items_with_errors = []
    for i, content in tqdm(enumerate(content_list), total=len(content_list)):
        original_name, fddb_name = content
        try:
            if fddb_name is None:
                informations = get_food_informations(content[0])
                food_items[i + 2] = informations
                time.sleep(random.randint(1, 10) / 10)
        except Exception as e:
            items_with_errors.append(original_name)
            logging.error("Error while parsing {}: {}".format(content, e))
    for row, infos in food_items.items():
        ws.range((row, 2)).value = infos["product_name"]
        ws.range((row, 3)).value = infos["serving_value"]
        ws.range((row, 4)).value = infos["serving_unit"]
        ws.range((row, 5)).value = infos["calories"]
        ws.range((row, 6)).value = infos["protein"]
        ws.range((row, 7)).value = infos["fat"]
        ws.range((row, 8)).value = infos["carbs"]
        ws.range((row, 9)).value = infos["relation"]
        ws.range((row, 10)).value = infos["relation_numerical"]
        ws.range((row, 11)).value = infos["product_link"]
    print("Finished crawling")
    print("These items could not be parsed: {}".format(items_with_errors))


if __name__ == "__main__":
    main()

