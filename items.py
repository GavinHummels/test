# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class MeItem(scrapy.Item):
	title = scrapy.Field()
	time = scrapy.Field()
	publisher = scrapy.Field()
	text = scrapy.Field()
