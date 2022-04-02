# scrape free legal working days for designating url

import json, requests, bs4
from datetime import date
month = {"ianuarie": 1, "februarie": 2, "martie": 3, "aprilie": 4, "mai": 5, "iunie": 6, "iulie": 7, "august": 8, "septembrie": 9, "octombrie": 10, "noiembrie": 11, "decembrie": 12}
SETTINGS = "settings.json" #rel path to settings

class Libere:
    def __init__(self, main_url, selector):
        self.date = date.today()
        self.main_url = main_url + str(self.date.year)
        self.selector = selector
        self.settings = Libere.loadJson(SETTINGS)

    def crawl(self):
        result = []
        res = requests.get(self.main_url)
        res.raise_for_status()
        selectAllData = self.selector
        page = bs4.BeautifulSoup(res.text, 'html.parser')
        for element in page.select(selectAllData):
            result.append(element.text)

        # processing data
        self.result = Libere.processing(result[1:], self.date.year)

    def write_to_file(self):
        with open(self.settings['output_folder'] + "libere_" + str(self.date.year) + ".txt", "w") as file:
            for item in self.result:
                file.write(item + "\n")

    @staticmethod
    def loadJson(name):
        with open(name, 'r') as file:
            data = json.loads(file.read())
        return data

    @staticmethod
    def processing(data, year): #array data
        result = []
        for item in data:
            word1 = item.split(' ')[0]
            word2 = item.split(' ')[1]
            word2 = month[word2]
            newWord = word1 + "/" + str(word2) + "/" + str(year)
            result.append(newWord)
        return result

if __name__ == '__main__':
    obj = Libere("https://www.zileliberelegale.ro/zile-libere-", "div.list-days-date")
    obj.crawl() ; obj.write_to_file()
