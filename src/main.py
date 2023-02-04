# This program scrapes the information from yazio.com to get calories and other informations about food items

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time


def main():
    base_url = "https://www.yazio.com/de/kalorientabelle"
    food_items = []


def get_food_informations(item: str):
    search_url = f"https://www.yazio.com/de/search?q={item}"
    search_page = requests.get(search_url)
    search_soup = BeautifulSoup(search_page.content, "html.parser")
    # find all paragraphs with the class "gs-title"
    search_results = search_soup.find_all("a")

if __name__ == "__main__":
    get_food_informations("Apfel")