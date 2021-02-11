from urllib.parse import quote
from bs4 import BeautifulSoup
import requests
import xlrd
import xlwt

Headers = {
    "Upgrade-Insecure-Requests": "1",
    "Connection": "keep-alive",
    "Cache-Control": "max-age=0",
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36',
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,la;q=0.7,pl;q=0.6",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"
}

class baiduSpider():
    def __init__(self):
        super().__init__()

    def load_excel(self):#读取excel

        self.df = xlrd.open_workbook('上市公司名单.xlsx')
        self.df2 = xlwt.Workbook()
        self.sheet1 = self.df.sheet_by_name('Sheet1')
        self.sheet2 = self.df2.add_sheet('Sheet1')
        self.num = 0
        self.middle = ['1','2']
        for i in range(1,self.sheet1.nrows):
            self.cell_value = self.sheet1.cell_value(i, 1)
            # 爬虫
            self.Baidu_Spider(self.cell_value)
            if self.sign == 1:#sign等于1则表示网页中有信息
                self.middle.append(self.cell_value)
        print("爬取到内存中，接下来往excel中写入...")
        for j in self.middle:
            self.sheet2.write(self.num, 0, j)
            self.num +=1
        print('完成')
        self.df2.save('输出名单.xls')

    def Baidu_Spider(self,phrase):
        keyword1 = ['供应链金融','供应链管理','供应链融资']
        self.sign = 0
        for k in keyword1:
            #quote将关键词转为URL的格式
            ExcelReadPhrase = quote(phrase, 'utf-8')
            LastKey = quote(k, 'utf-8')
            try:
                url = 'https://www.baidu.com/s?wd=%s%s' % (ExcelReadPhrase,LastKey)
                response = requests.get(url, headers=Headers)
                soup = BeautifulSoup(response.text, 'html.parser')
                finaltext = soup.find('div', {'id': 'content_left'})
                if phrase in str(finaltext) and k in str(finaltext):
                    self.sign = 1
            except:
                print(url)
def main():
    begin = baiduSpider()
    begin.load_excel()

if __name__ == '__main__':
    main()
