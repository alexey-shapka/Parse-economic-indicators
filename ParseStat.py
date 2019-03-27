from dateutil.rrule import rrule, MONTHLY
from urllib.request import Request, urlopen
import bs4 as bs
import requests
import datetime
import xlwt
import time

class Parser:

    def __init__(self, years, filename):
        self.SalaryDict = {}
        self.USD = []
        self.EURO = []
        self.years = years
        self.dates = tuple(dt.strftime("%Y%m%d") for dt in rrule(MONTHLY, dtstart=datetime.date(min(years),1,1), until=datetime.date(max(years),12,1)))
        self.file = xlwt.Workbook(encoding="utf-8")
        self.first_sheet = self.file.add_sheet("First page", cell_overwrite_ok=True)
        self.filename = filename
        self.start_time = 0
        self.UpdateSalaryDict()
        self.UpdateCurrencyDicts()
        self.FillOutputFile()
        self.SaveFileAndSetWidthColomn()

    def UpdateSalaryDict(self):
        self.start_time = time.time()
        for y in self.years:
            self.req = Request('https://index.minfin.com.ua/labour/salary/average/{}'.format(str(y)), headers={'User-Agent': 'Mozilla/5.0'})
            self.webpage = urlopen(self.req).read()
            soup = bs.BeautifulSoup(self.webpage,'lxml')
            self.tables = soup.find_all("div", class_="glue-table")
            for i in self.tables:
                self.row = i.find_all('tr')
                for j in self.row:
                    self.values = list(map(lambda x: x.text, j.find_all('td')))
                    if len(self.values) != 0 and self.values[0] != 'г.Севастополь' and self.values[0] != 'АР Крым':
                        if self.values[0] not in self.SalaryDict:
                            self.SalaryDict[self.values[0]] = []
                            self.SalaryDict[self.values[0]] += self.values[1:]
                        else:
                            self.SalaryDict[self.values[0]] += self.values[1:]

    def UpdateCurrencyDicts(self):
        for i in self.dates:
            self.response_eur = requests.get('https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?valcode=EUR&date={}&json'.format(i))
            self.data_eur = self.response_eur.json()
            self.EURO.append(self.data_eur[0]['rate'])
            self.response_usd = requests.get('https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?valcode=USD&date={}&json'.format(i))
            self.data_usd = self.response_usd.json()
            self.USD.append(self.data_usd[0]['rate'])

    def FillOutputFile(self):
        self.regions = [*self.SalaryDict]
        
        for i in range(len(self.regions)):
            self.first_sheet.write(0, i+1, self.regions[i])
            for j in range(len(self.dates)):
                self.first_sheet.write(j+1, i+1, self.SalaryDict[self.regions[i]][j])

        self.first_sheet.write(0, len(self.regions)+1, 'USD')
        self.first_sheet.write(0, len(self.regions)+2, 'EURO')
        for i in range(len(self.dates)):
            self.first_sheet.write(i+1, 0, '{}-{}-{}'.format(self.dates[i][:4], self.dates[i][4:6], self.dates[i][6:]))
            self.first_sheet.write(i+1, len(self.regions)+1, self.USD[i])
            self.first_sheet.write(i+1, len(self.regions)+2, self.EURO[i])

    def SaveFileAndSetWidthColomn(self):
        for i in range(len(self.SalaryDict)+3):
            self.first_sheet.col(i).width = 256 * 20
        self.file.save("{}.xls".format(self.filename))

        print('Succsesfully saved.\nRunning Time: {}'.format(time.time() - self.start_time))

p = Parser((2010,2011,2012,2013,2014), "Newfile")