from scrapy.cmdline import execute

def main():
    # execute(['scrapy', 'crawl', 'FARA_scraper', '--nolog'])
    execute(['scrapy', 'crawl', 'otkritie'])

if __name__ == '__main__':
    main()