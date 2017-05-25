# -*- coding: UTF-8 -*-

import xlsxwriter
import sys
import validators
import os
import urllib
import Image
import json
import re

thumbSize = 128, 128

inFile = sys.argv[1]
print inFile
outFile = inFile.split('.')[0] + '.xlsx'
print outFile
picDir = inFile.split('.')[0]
try:
	os.stat(picDir)
except:
	os.mkdir(picDir)

workbook = xlsxwriter.Workbook(outFile)
worksheet = workbook.add_worksheet()


def picDownload(index,url):
	print "url pic: " + url
	if validators.url(url):
		outFile = picDir + '/' + str(index) + ".jpg"
		print url.split('.')[-1].strip('"')[0:-1]
		pic = open( outFile,'wb')
		pic.write(urllib.urlopen(url).read())
		pic.close()
		makeThumb(outFile)
		return True
	else:
		return False

def makeThumb(inFile):
	outFile = os.path.splitext(inFile)[0] + ".thumbnail"
	if inFile != outFile:
		try:
            		im = Image.open(inFile)
            		im.thumbnail(thumbSize, Image.ANTIALIAS)
            		im.save(outFile, "JPEG")
        	except IOError:
            		print "cannot create thumbnail for '%s'" % inFile


def sizing(size,nosize):
	print "len size:" + str(len(size))
	print "len nosize: " + str(len(nosize))
	size_d = {}
	size_l = []
	size_ls = []
	size = size.lstrip('"').strip('"').strip('\n').split(',')
	for i in size:
		#print "REEEEE " + re.sub("^\s+|\n|\r|\t|\s+$", '',i)
		#print i.replace('""','"' ,100).replace("[",'',1).replace("]\"","",1).replace("]","",1).replace('\t',"",1000).replace('\n',"",1000)
		size_l.append(json.loads(i.replace('""','"' ,100).replace("[",'',1).replace("]\"","",1).replace("]","",1).replace('\t',"",1000).replace('\n',"",1000)))
	for i in  size_l:
		size_ls.append(i['size'])


	nosize_l = []
	nosize_ls = []
	nosize = nosize.lstrip('"').strip('"').strip('\n').split(',')
	print "len nosize: " + str(len(nosize))
	if len(nosize) > 1:
		for i in nosize:
			nosize_l.append(json.loads(i.replace('""','"' ,100).replace("[",'',1).replace("]\"","",1).replace("]","",1).replace('\t',"",1000).replace('\n',"",1000)))
		for i in  nosize_l:
			nosize_ls.append(i['nosize'])
	else:
		print 'im here'
		nosize_ls = []
	for i in nosize_ls:
		size_ls.remove(i)
	size_ls =  list(set(size_ls))
	print nosize_ls
	print size_ls
	return size_ls



		
#	print "size: " + size.strip('"')[0]
#	print "nosize: " + str(nosize)
#	for i in size.strip('"'):
#		print i
#		#size_l.append(i['size'])
#	for i in nosize:
#		#nosize_l.append(i['nosize'])
#	for i in nosize_l:
#		size_l.remove(i)
#	print size_l
	



with open(inFile) as f:
	colsNames = f.readline().strip('\n').split(',')
	print colsNames
	c = 0 
	for col in colsNames:
		worksheet.write(0, c, unicode(col, 'utf-8'))
		c += 1
	picIndex = colsNames.index('pic-src')
	colSize = colsNames.index('size')
	colNoSize = colsNames.index('nosize')
	rowId = 1
	for line in f:
		cols = line.split('",')
		print "index: " + str(rowId)
		sizes = sizing(cols[colSize],cols[colNoSize])
		if picDownload(rowId,cols[picIndex].strip('"')):
			print "download photo"
		else:
			print "no photo"
		colId = 0
		for col in cols:
			if colId == picIndex:
				#worksheet.insert_image(rowId, colId, picDir + '/' + str(rowId) + cols[picIndex].strip('"')[-3::1], {'x_scale': 0.050, 'y_scale': 0.050})
				#worksheet.insert_image(rowId, colId, picDir + '/' + str(rowId) + ".thumbnail", {'x_scale': 0.050, 'y_scale': 0.050})
				worksheet.insert_image(rowId, colId, picDir + '/' + str(rowId) + ".thumbnail") 
			elif colId == colSize:
				worksheet.write(rowId, colId, ','.join(sizes))
			else:
				worksheet.write(rowId, colId, unicode(col.strip('"'), 'utf-8'))
			print str(rowId) + ' ' + str(colId) + ": " + col
			colId += 1
		rowId += 1

workbook.close()


