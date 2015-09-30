# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class VonawebItem(scrapy.Item):
    # define the fields for your item here like:
    # name = scrapy.Field()
    name = scrapy.Field()
    catalog = scrapy.Field()
    image_urls = scrapy.Field()
    images = scrapy.Field()
    pack = scrapy.Field()
    index = scrapy.Field()
    brand = scrapy.Field()

