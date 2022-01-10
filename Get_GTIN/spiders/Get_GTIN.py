# -*- coding: utf-8 -*-
import scrapy
from scrapy.exceptions import CloseSpider
from scrapy import Request
from urlparse import urlparse
from json import loads
from datetime import date
from collections import OrderedDict
# import json
import re, time
from xlrd import open_workbook

from openpyxl import Workbook, load_workbook
import numpy

class Get_GTINSpider(scrapy.Spider):

    name = "Get_GTIN_spider"

    use_selenium = False

    next_count = 1
    gtin_data_list = []
    key_list = []

    default_seller_names = ['carethy.es','naturaldietetic.com', 'naturitas', 'mifarma.es', 'dieteticacentral.com']
###########################################################

    def __init__(self, *args, **kwargs):
        super(Get_GTINSpider, self).__init__(*args, **kwargs)

        self.supplier_reference_list = []


        wb = open_workbook('Categories Managment.xlsx')
        for s in wb.sheets():
            #print 'Sheet:',s.name


            if s.name == 'Categories Managment':
                for row in range(s.nrows):
                    if row == 0:
                        continue
                    for col in range(s.ncols):
                        if col == 0:
                            value = s.cell(row, col).value
                            self.supplier_reference_list.append(value)
        # del wb

###########################################################

    def start_requests(self):
        i = 0
        for supplier_reference in self.supplier_reference_list:
            i += 1
            if i < 3000:
                time.sleep(1)
                url = 'https://www.google.es/search?biw=1680&bih=944&tbm=shop&ei=O97mWujjF8T1gQawooOoBw&q={0}&oq={0}'.format(supplier_reference)
                yield Request(url, callback=self.parse, dont_filter=False, meta={'key': supplier_reference})


        ### testhttps://www.google.com/search?biw=907&bih=944&tbm=shop&ei=j8LgWoPhEoKsjwPL4bTABg&q={}&oq={}
        # url = 'https://www.google.es/search?biw=1680&bih=944&tbm=shop&ei=O97mWujjF8T1gQawooOoBw&q={0}&oq={0}'.format('SAIS-005')
        # yield Request(url, callback=self.parse, dont_filter=False, meta={'key': 'ART-0443'})

###########################################################

    def parse(self, response):
        print response.url

        href = response.xpath('.//*[@class="sh-dlr__thumbnail"]/a/@href').extract_first()
        if href:
            url = 'https://www.google.es' + href
            yield Request(url, callback=self.parseGTIN, dont_filter=False, meta=response.meta)

    def parseGTIN(self, response):

        data = response.xpath('.//*[@class="section-inner"]/div/div/span/text()').extract()
        seller_names = response.xpath('.//*[@class="os-seller-name-primary"]/a/text()').extract()
        i = 0
        gtin_data = ''
        for d in data:
            if d == 'GTIN':
                for seller_name in seller_names:
                    if self.default_seller_names.__contains__(seller_name.lower()):
                        gtin_data = data[i + 1]
                        break
            i += 1
        try:
            gtin_data = str(int(gtin_data))
            self.gtin_data_list.append(gtin_data)
        except:
            self.gtin_data_list.append('')
        self.key_list.append(response.meta['key'])

        print {'key': response.meta['key'], 'gtin_data': gtin_data}

        # yield {'gtin': gtin_data}
        #
        #
        # test_str = '<script type="text/javascript"> var WCParamJS = {"storeId":"10001","catalogId":"10001","langId":"-1000","pageView":"grid","orderBy":"","orderByContent":"","searchTerm":""}; var absoluteURL = "https://www.abcdin.cl/tienda/"; var imageDirectoryPath = "/wcsstore/ABCDIN/"; var styleDirectoryPath = "images/colors/color1/";	var supportPaymentTypePromotions = true;'
        #
        #
        # s = re.findall('var WCParamJS = (\{.*\};)',response.body.replace('\n', ''))[0].split('};')[0] + '}'
        #
        # a = loads(s.replace('\t', '').replace("'", '"'))
        #
        # print len(products)
        # print "###########"
        # print response.url
        #
        # if not products: return
        #
        # for i in products:
        #     item = OrderedDict()
        #
        #     item['Vendedor'] = 140
        #
        #     detail_url = i.xpath('.//*[@class="product_name"]/a/@href').extract_first()
        #     if not detail_url:
        #         continue
        #     item['ID'] = detail_url.split('-')[-1]
        #     item['Title'] = i.xpath('.//*[@class="product_name"]/a/text()').extract_first().strip()
        #
        #     price = i.xpath('.//*[contains(@id,"offerPrice_")]/text()').extract()#.re(r'[\d.,]+')
        #     if price:
        #         price = price[1].split('$')[-1].strip().replace('.', '')
        #     else:
        #         price = i.xpath('.//*[contains(@id,"listPrice_")]/text()').extract()#.re(r'[\d.,]+')
        #         if price:
        #             price = price[1].split('$')[-1].strip().replace('.', '')
        #         else:
        #             continue
        #     item['Price'] = price
        #     item['Currency'] = 'CLP'
        #
        #     item['Category URL'] = 'https://www.abcdin.cl' + response.meta['CatURL']
        #     item['Details URL'] = detail_url
        #     item['Date'] = str(date.today())
        #
        # # count = len(response.xpath('//ul[@class="pagination"]/li'))
        # next00 = response.xpath('//*[@class="right_arrow "]')
        #
        #
        # if next00:
        #     response.meta['page_num'] += 1
        #     next_url = 'https://www.abcdin.cl' + response.meta['CatURL'] + '#facet:&productBeginIndex:{0}&orderBy:&pageView:grid&minPrice:&maxPrice:&pageSize:&'.format(response.meta['page_num'] * 24)
        #
        #     yield Request(next_url, callback=self.parse, meta=response.meta, dont_filter=True)
        