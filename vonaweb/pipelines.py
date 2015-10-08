# -*- coding: utf-8 -*-
import scrapy
from scrapy.contrib.pipeline.images import ImagesPipeline
import json
import os
# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html


class VonawebPipeline(ImagesPipeline):

    def get_media_requests(self,item,info):
        for image_url in item['images_urls']:
            yield scrapy.Request(image_url)

    def item_completed(self,results,items,info):
        image_paths = [x['path'] for ok,x in results if ok]
        if not image_paths:
            item['image_paths']=None
        else:
            item['image_paths'] = image_paths
        return item

class MyItemPiple(object):
    def __init__(self):
        self.file=open('vona.json','wb')
        self.list=[]

    def process_item(self,item,spider):
#        line = json.dumps(dict(item)) 
#        self.file.write(line+"\r\n")
        self.list.append(dict(item))
        return item

    def close_spider(self,spider):
        jsondata = json.dumps(self.list)
        self.file.write(jsondata)
        self.file.close()
