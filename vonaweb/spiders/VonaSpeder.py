#coding=utf-8
import xlrd
import scrapy
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
from scrapy.http.request import Request
from scrapy.contrib.spiders import Rule
from scrapy.selector import Selector
from vonaweb.items import VonawebItem
import re
import os

class VonaSpider(scrapy.Spider):
    name = 'vona'
    allowed_domains = ['th.stg.misumi-ec.com']
    start_urls = []

    def __init__(self):
        prefix = os.getcwd()
        print prefix
        data = xlrd.open_workbook(prefix+os.path.sep+'file'+os.path.sep+u'28800-28999逸逵.xls')
        table = data.sheets()[0]
        for url in table.col_values(3)[2:]:
            self.start_urls.append(url)

    def start_requests(self):
        for i,url in enumerate(self.start_urls):
            print i,url
            yield scrapy.FormRequest(url,cookies={'passcode':'vonaweb'},meta={'index':i},callback=self.parse_item)

    def after_login(self,response):
        if 'パスワード' in response:
            self.logger.error('Login failed')
            print 'Login failed'
            return
        else :
            self.logger.info('Login success')


    def parse_item(self,response):
        sel = Selector(response)
        name = sel.xpath('//h1[contains(@itemprop,"name")]//text()').extract()
        catalog = sel.xpath('//a[contains(@id,"Tab_catalog")]/text()').extract()
        image_url = sel.xpath('//img[contains(@itemprop,"image")]/@src').extract()
        brand = sel.xpath('//div[contains(@class,"brand")]/span/a[contains(@itemprop,"brand")]/span[contains(@itemprop,"name")]/text()').extract()
        index = response.meta['index']
        item = VonawebItem(index=index,pack=False)
        pattern = re.compile(ur'.*\u3010.*Pieces Per Package.*\u3011')
        if name:
            item['name']=u''
            for each in name:
                item['name'] += each 
        if catalog:
            item['catalog'] = catalog[0].encode('utf-8')
        else :
            item['catalog'] = None 
        if name:
            match = pattern.match(name[0])
            if match:
                item['pack']=True
        else:
            item['pack']=False
        if image_url:
            item['image_urls']=['http://th.stg.misumi-ec.com'+image_url[0]]
        if brand:
            item['brand']=brand[0].encode('utf-8')

        return item
