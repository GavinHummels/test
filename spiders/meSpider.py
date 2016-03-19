#coding:utf-8
import scrapy
from scrapy.http import Request
from me.items import MeItem




class meSpider(scrapy.Spider):
	name = "mespider"
	allowed_domains = ["http://me.seu.edu.cn/"]
	start_urls = ["http://me.seu.edu.cn/1550/list.htm",#学院新闻
	"http://me.seu.edu.cn/1549/list.htm",#通知公告
	"http://me.seu.edu.cn/1323/list.htm",#本科生教务
	"http://me.seu.edu.cn/1553/list.htm",#学生工作
	"http://me.seu.edu.cn/1333/list.htm",#研究生教务
	"http://me.seu.edu.cn/1551/list.htm"]#学术论坛

	def parse(self,response):
		contents=response.xpath("//div[@id='wp_news_w3']//table//tr")
		item=MeItem()
		for content in contents:
			try:
				title=content.xpath("./td[1]/a/@title").extract()[0].encode("gbk")
				time=content.xpath("./td[2]/div/text()").extract()[0].encode("gbk")
				publisher=content.xpath("./td[3]/div/text()").extract()[0].encode("gbk")
				item["title"]=title
				item["time"]=time
				item["publisher"]=publisher
				yield item
			except:
				pass
			