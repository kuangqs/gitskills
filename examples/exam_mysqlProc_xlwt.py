# -*- coding:utf-8 -*- 

import xlwt 
import	MySQLdb
import	datetime
import	sys
import	string
reload(sys)
sys.setdefaultencoding('utf-8')

## 执行存储过程 参考：http://www.cnblogs.com/luoshulin/archive/2009/10/28/1591385.html
def	callproc(filename, proc):
	
	conn=MySQLdb.connect(host='127.0.0.1',port=3306,user='xingyanshi',passwd='123456',db='report',charset='utf8')
	cur=conn.cursor()
	cur.nextset()
	#cur.execute('call	actioninfo(%s,%s,%s,%s,%s,%s, @odate, @ogame_channel, @oagent, @oAdCreative, @click_num, @action_num, @action_new_num, @ad_au_5_num )',(gameid,channel,agent,startdate,enddate,platform))
	cur.execute(proc)

	data=cur.fetchall()
	#print data

	fields = cur.description
	print fields
	cur.close()
	conn.close()
	
	header =[]
	for field in range(0,len(fields)):
		header.append(fields[field][0])
	#print header

	write_excel(filename, header, data)
	
	return	data

## 写入excel 
## 参考：
## http://www.jb51.net/article/60510.htm
## http://www.jb51.net/article/77626.htm
## 后边可以研究一下怎么合并xls文件，将结果写到一个文件里面去
def write_excel(filename, header, data):
	
	workbook = xlwt.Workbook(encoding = 'utf-8')
	sheet = workbook.add_sheet('data',cell_overwrite_ok=True)

	# 写上头字段信息
	for h in range(0,len(header)):
		sheet.write(0, h, header[h])

	# 获取并写入数据段信息
	row = 1
	col = 0
	for row in range(1,len(data)+1):
		for col in range(0,len(header)):
			sheet.write(row,col,u'%s'%data[row-1][col])
			#print u'%s'%data[row-1][col],
		#print ''

	print "please check %s" % filename
	workbook.save(filename)
	
	return filename
	
## Usage
def Usage(pyname):
	print "Usage: python %s gameid channel agent startdate enddate platform" % pyname
	print '''Example: python %s 1479458217005 万通 万通 2017-06-06 2017-07-28 ios正版 ''' % pyname
	sys.exit(255)
	

def main():

	## 传参数 更好方法参考：http://www.360doc.com/content/16/0424/16/31913486_553405472.shtml
	#print sys.argv
	sys.argv = [ a.decode('gbk') for a in sys.argv]
	#print sys.argv
	if len(sys.argv) == 7:
		pass
	else:
		Usage(sys.argv[0])
	
	gameid=sys.argv[1]
	channel=sys.argv[2]
	agent=sys.argv[3]
	startdate=sys.argv[4]
	enddate=sys.argv[5]
	platform=sys.argv[6]
	
	## 存储过程字典:k为xls文件名的头，value为要执行的存储过程
	proc_contents = {
		### -- 01.spend
		'spend':u'''CALL spend('%s', '%s', '%s', '%s', '%s', @oplatform , @oclass_ad , @oclass_A)''' % (gameid,channel,agent,startdate,enddate),
		### -- 02. actioninfo (激活汇总 sheet)
		'actioninfo':u'''call actioninfo('%s', '%s', '%s', '%s', '%s', '%s', @odate, @ogame_channel, @oagent, @oAdCreative, @click_num, @action_num, @action_new_num, @ad_au_5_num )''' % (gameid,channel,agent,startdate,enddate,platform),
		### -- 03.day_recharge (每日充值情况 sheet)
		'day_recharge':u'''CALL day_recharge('%s', '%s', '%s', '%s', '%s', '%s', @odate , @omoney)''' % (gameid,channel,agent,startdate,enddate,platform),
		### 4/5/6...待补充
	}
	
	## xls文件名头
	dt = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
	
	for k,v in proc_contents.items():
		print k,":"
		filename="%s_%s.xls" % (k, dt)
		callproc(filename,v)


if __name__ == '__main__':
	main()
