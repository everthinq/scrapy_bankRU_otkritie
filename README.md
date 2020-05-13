Please, use UNIX based operation system to test this solution. Everything was written for UNIX based OS. When I tried to test it on Windows, I had couple of problems with installing Scrapy framework (C++14 error, etc). 

I'm using Python3 and Scrapy framework as core, openpyxl as excel sheets reader/writer

### This scrapy project parses [Otkritie Bank of Russia map](https://www.open.ru/map) through API to get branches data and stores scraped data into *.xlsx file

### Instructions:

0) Install Python3 -- https://www.python.org/
1) Install virtualenv -- *pip install virtualenv*
2) Create a virtual environment inside that folder -- *virtualenv venv*
3) Activate the virtual environment -- *source venv/bin/activate*
4) Install requirements.txt -- *pip install -r requirements.txt*
5) To open the main source file, go to *scrapy_otkritie/scrapy_otkritie/spiders -- otrkitie.py*
6) You can run spider through terminal command -- *scrapy crawl otkritie*   
6.5) Or you can run main.py which does the same thing -- *scrapy crawl otkritie*
