# -*- coding:utf-8 -*-
import urllib2
import re
import socket
import xlrd
import xlwt

#解决中文显示
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

#写入excel
def writexlsxdate(sheeti,rowlinenumi,rowlinedate):
    for rowwi in range(0,len(rowlinedate)):
        sheeti.write(rowlinenumi,rowwi,rowlinedate[rowwi])

outputfile = xlwt.Workbook()#写文件
sheet1 = outputfile.add_sheet(u'sheet1',cell_overwrite_ok=True)#初始化sheet
filedir = raw_input("输入WAF导出的excel文件:")
filedir = unicode(filedir , "utf8")
outputdir = raw_input("保存过滤后的excel文件:")
outputdir = unicode(outputdir,"utf8")
data = xlrd.open_workbook(filedir)#打开文件地址
table = data.sheets()[0]#读取文件内容
nrows = table.nrows#行化
nofindpage=0#初始化404统计
filtersumline=0#初始化过滤统计
nofiltersumline=0#舒适化未过滤统计
mm = False#初始化过滤bool
rowlinei = 0 #初始化写文件行号

filterstr = raw_input("请输入需要过滤的关键字使用 , 隔开:")#输入过滤单词
filterbaby=filterstr.split(',')#分割过滤单词
print "要过滤的关键字为:%s"%filterbaby#输出过滤词
print "过滤的关键字数为:%s" %len(filterbaby)#统计过滤词数量

#以下主体过滤单词、访问url匹配相应码
for rowsline in range(16,nrows,1):
	writerowdate = table.row_values(rowsline)
#	print writerowdate
	fdurl = table.cell(rowsline,5).value
	if len(fdurl) > 180:
#		print "len is:%s"%len(fdurl)
#		refdurl = re.search(r"^[a-z,A-Z,0-9,\,/,.,:,-,_]+",fdurl)
#		fdurl = refdurl.group(0)[0:180]
		fdurl = re.sub(r'\?$','',fdurl)#过滤结尾的‘？’
	else:
		fdurl = fdurl[0:180]
		fdurl = re.sub(r'\?$','',fdurl)#过滤结尾的‘？’
#	print fdurl
	for babygo in filterbaby:
		mm = bool(mm) or bool(re.search(babygo,fdurl))#通过bool判断是否存在过滤词
	if not mm:
		mm = False
		try:
			nofiltersumline=nofiltersumline+1
			d = urllib2.urlopen('http://'+fdurl+'',timeout = 5)
			if d.code == 404:
				print "%s %s------>网页状态码为 %s"%(rowsline,fdurl,d.code)
				nofindpage=nofindpage+1
			else: 
				print "%i %s------>网页状态码为 %s"%(rowsline,fdurl,d.code)
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1
		except UnicodeError:
			print "%s %s------>unicodeerror"%(rowsline,fdurl)
			writexlsxdate(sheet1,rowlinei,writerowdate)
			rowlinei =rowlinei + 1
		except socket.error,socketerror:				
			print "%s %s------>网络错误 "%(rowsline,fdurl)+"%s"%socketerror
			writexlsxdate(sheet1,rowlinei,writerowdate)
			rowlinei =rowlinei + 1
   		except urllib2.HTTPError,e:
#			try:
			if e.code == 404:
				print "%s %s---->网页状态码为!!!!404"%(rowsline,fdurl)
				nofindpage=nofindpage+1
			else:
				print "%s %s------>网页状态码为"%(rowsline,fdurl)+"%s"%e.code
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1

#			except:
#				print "%s------>other error"%fdurl
#				writexlsxdate(sheet1,rowlinei,writerowdate)
#				rowlinei =rowlinei + 1
   		except urllib2.URLError,e:
			if re.search('10061',str(e.reason)):
				print "%s %s------>"%(rowsline,fdurl)+"被远程主动拒绝（rst）"
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1
			elif re.search('11001',str(e.reason)):
				print "%s %s------>"%(rowsline,fdurl)+"未查找到域名"
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1
			else:
				print "%s %s------>网络错误 %s"%(rowsline,fdurl,e.reason)
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1
		except:
				print "%s------>other error"%fdurl
				writexlsxdate(sheet1,rowlinei,writerowdate)
				rowlinei =rowlinei + 1
	else:
		filtersumline=filtersumline+1
		mm = False
print "code 404 %s"%nofindpage
print "filtered %s"%filtersumline
print "find %s"%nofiltersumline

outputfile.save(outputdir)

#.asp,.php,.txt,.cfg,.cfc,.gif

