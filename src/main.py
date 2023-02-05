# This program scrapes the information from yazio.com to get calories and other informations about food items

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

BASE_URL = "https://fddb.info"
SEARCH_URL = "https://fddb.info/db/de/suche/?udd=0&cat=site-de&search={}"


def main():
    food_items = []


def get_food_informations(item: str):
    search_url = SEARCH_URL.format(item)
    search_page = requests.get(search_url)
    search_soup = BeautifulSoup(search_page.content, "html.parser")
    # find the first result from a table which has a link with "/lebensmittel/" in it
    tables = search_soup.find_all("table")
    product_link = None
    for table in tables:
        product_link = table.find("a", href=lambda href: href and "/db/de/lebensmittel/" in href)
        if product_link is not None:
            product_link = BASE_URL + product_link.attrs["href"]
            break
    if product_link is not None:
        product_html = requests.get(product_link)
        product_soup = BeautifulSoup(product_html.content, "html.parser")
        # subset product_soup to only the div with the class standardcontent
        calories_table = product_soup.find("div", class_="standardcontent")
        # get the value of the div after the div with the value "Kalorien"
        calories = calories_table.find("div", text="Kalorien").find_next("div").string
        protein = calories_table.find("div", text="Protein").find_next("div").string
        fat = calories_table.find("div", text="Fett").find_next("div").string
        carbs = calories_table.find("div", text="Kohlenhydrate").find_next("div").string
        # get string of the div witht the class "itemsec2012"
        relation = calories_table.find("div", class_="itemsec2012").string

        # subset the product_soup to only the dev with the class "rightblock"
        right_block = product_soup.find("div", class_="rightblue-complete")
        # get from right block the string from the first div with class "serva"
        serving_size_text = right_block.find("div", class_="serva").find("a").string
        serving_size = serving_size_text.split("(")[-1][:-1]
        value, unit = serving_size.split(" ")
        # get the string from the h1 heading with id "fddb-headline1"
        product_name = product_soup.find("h1", id="fddb-headline1").string
        return {
            "product_name": product_name,
            "serving_size_text": serving_size_text,
            "serving_value": value,
            "serving_unit": unit,
            "calories": calories,
            "protein": protein,
            "fat": fat,
            "carbs": carbs,
            "relation": relation,
            "product_link": product_link,
        }


if __name__ == "__main__":
    get_food_informations("Apfel")