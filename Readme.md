# FoodCrawler

This is a simple excel combined with a small python application which I developed to track my recipes and the 
calories/macros. The vba code parses the ingredients from the cells and creates lookup functions while the python
code just scrapes the informations from the configured website. 

So far I have only implemented a function for the german website fddb.info, but others can be integrated as well.

## Usage

I have developed and tested this under Windows. Other OS might work as well, but I can't guarantee it.
At least you need to compile the python program yourself.



1. Download the excel file and enable macros

For Windows OS:
2. Go to the "Overview" tab and click "Download Python Program" (use another path if you want, Cell B2)

For other OS run the following:
2a. Run the following commands
```` shell
git clone git@github.com:FelixKleineBoesing/CaloriesCrawler.git
cd CaloriesCrawler
pip install -r requirements.txt
pip install -r requirements_dev.txt
pyinstaller  -F foodcrawler/main.py
````
2b. Copy the executable from the dist folder to the location which is defined in the excel overview tab (Cell B2)

3. Track your recipes in the tabs. Feel free to create new tabs or remove them. A worksheet needs to have the 
WorkSheet_Change sub which gets triggered, when you change a value in the column D.
4. If you have tracked your recipes and want to download the Ingredient data go to overview and click on "Run INgredients Updating".
This is not done on the fly, because it takes a while and I don't want to slow down the excel.
