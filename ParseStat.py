from dateutil.rrule import rrule, MONTHLY
from urllib.request import Request, urlopen
import tradingeconomics as te
import bs4 as bs
import requests
import datetime
import xlwt
import time

class Parser:

    def __init__(self, years, filename):
        self.years = years
        self.cars = ('toyota', 'renault', 'volkswagen')
        self.GlobalData = {self.cars[0] : [], self.cars[1] : [], self.cars[2] : [], 'USD' : [], 'EURO'  : [], 'Inflation' : [], 'RWI' : []}
        self.dates = tuple(dt.strftime("%Y%m%d") for dt in rrule(MONTHLY, dtstart=datetime.date(min(years),1,1), until=datetime.date(max(years),12,1)))
        self.file = xlwt.Workbook(encoding="utf-8")
        self.first_sheet = self.file.add_sheet("First page", cell_overwrite_ok=True)
        self.filename = filename
        self.start_time = 0
        self.GetSalaryStatistics()
        self.GetCarSellRating()
        self.GetCurrencyStatistics()
        self.GetInflationStatistics()
        self.GetRealWagesIndex()
        self.FillAndSaveFile()

    def GetSalaryStatistics(self):
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
                        if self.values[0] not in self.GlobalData:
                            self.GlobalData[self.values[0]] = []
                            self.GlobalData[self.values[0]] += self.values[1:]
                        else:
                            self.GlobalData[self.values[0]] += self.values[1:]

    def GetRealWagesIndex(self):
        self.req = Request('https://index.minfin.com.ua/labour/salary/index/', headers={'User-Agent': 'Mozilla/5.0'})
        self.webpage = urlopen(self.req).read()
        soup = bs.BeautifulSoup(self.webpage,'lxml')
        self.tables = soup.find("div", class_="idx-block-1320 compact-table").find_all('tr')[1:]
        for i in self.tables:
            self.year_header = int(i.find('th').text)
            if self.year_header in self.years:
                self.GlobalData['RWI'] += list(map(lambda x: x.text, i.find_all('td')))

    def GetInflationStatistics(self):
        self.req = Request('https://index.minfin.com.ua/ua/economy/index/inflation/', headers={'User-Agent': 'Mozilla/5.0'})
        self.webpage = urlopen(self.req).read()
        soup = bs.BeautifulSoup(self.webpage,'lxml')
        self.tables = soup.find("div", class_="idx-block-1320 compact-table").find_all('tr')[1:]
        for i in self.tables:
            self.year_header = int(i.find('th').text)
            if self.year_header in self.years:
                self.GlobalData['Inflation'] += list(map(lambda x: x.text, i.find_all('td')))

    def GetCarSellRating(self):
        for y in self.years:
            for car in self.cars:
                self.req = Request('https://auto.vercity.ru/statistics/sales/europe/{}/ukraine/{}/'.format(str(y), car), headers={'User-Agent': 'Mozilla/5.0'})
                self.webpage = urlopen(self.req).read()
                soup = bs.BeautifulSoup(self.webpage,'lxml')
                self.data = list(map(lambda x: x.text, soup.find("table", class_="page_brands").findNext('tbody').findNext('tr').find_all('td')))[3:-2]
                self.GlobalData[car] += self.data
                
    def GetCurrencyStatistics(self):
        for i in self.dates:
            self.response_eur = requests.get('https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?valcode=EUR&date={}&json'.format(i))
            self.data_eur = self.response_eur.json()
            self.GlobalData['EURO'].append(self.data_eur[0]['rate'])
            self.response_usd = requests.get('https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?valcode=USD&date={}&json'.format(i))
            self.data_usd = self.response_usd.json()
            self.GlobalData['USD'].append(self.data_usd[0]['rate'])

    def FillAndSaveFile(self):
        self.titles = [*self.GlobalData]

        self.first_sheet.col(0).width = 256 * 20
        for i in range(1, len(self.titles)+1):
            self.first_sheet.write(0, i, self.titles[i-1])
            self.first_sheet.col(i).width = 256 * 20

        for i in range(len(self.dates)):
            self.first_sheet.write(i+1, 0, '{}-{}-{}'.format(self.dates[i][:4], self.dates[i][4:6], self.dates[i][6:]))
            for j in range(len(self.GlobalData)):
                 self.first_sheet.write(i+1, j+1, self.GlobalData[self.titles[j]][i])

        self.file.save("{}.xls".format(self.filename))
        print('Succsesfully saved.\nRunning Time: {}'.format(time.time() - self.start_time))

p = Parser((2011,2012,2013,2014,2015,2016,2017,2018), "Parse(2011-2018)")