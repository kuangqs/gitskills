# -*- coding:utf-8 -*- 

'''
sshtunnel MySQLdb xlwt
设置ssh代理，访问mysql数据库，查询出结果保存到xls
'''

import xlwt 
import MySQLdb
import datetime
from sshtunnel import SSHTunnelForwarder
import sys   
reload(sys) # Python2.5 初始化后会删除 sys.setdefaultencoding 这个方法，我们需要重新载入   
sys.setdefaultencoding('utf-8')   

with SSHTunnelForwarder(
         ('127.0.0.1', 22),    #B机器的配置
         ssh_password="passsword",
         ssh_username="xingyanshi",
         remote_bind_address=('127.0.0.2', 3306)) as server:  #A机器的配置

    conn = MySQLdb.connect(host='127.0.0.1',              #此处必须是是127.0.0.1
                           port=server.local_bind_port,
                           user='xingyanshi',
                           passwd='password',
                           db='user')


    user_id_list='''123456,234567,345678'''

    end_pay_time='''2017-8-4'''

    sql = '''select us.truename as 姓名, bi.user_id as 用户ID,bi.order_id as 订单ID , bi.id as 账单ID, bf.`trade_no` as 交易号,bf.pay_time as 还款时间, bi.`capital` as 本金, bi.`interest` as 利息, bf.amount as 实际还款,(bi.paid-bi.`capital`) as 需报销金额 
from bill_items bi 
join bill_item_paid_flows bf on bi.id=bf.bill_item_id 
join users us on bi.user_id=us.id
where bi.user_id in (%s) and bi.pay_status=2 and bi.status=1 and bi.pay_time>='2017-6-12' and bi.pay_time<'%s' and (bi.paid>bi.`capital`) 
order by bi.user_id, bi.order_id,bi.id; ''' % (user_id_list, end_pay_time)

    cur=conn.cursor()
    cur.execute(sql)
    results = cur.fetchall()
    #print results 
    fields = cur.description

    cur.close()
    conn.close()

    workbook = xlwt.Workbook(encoding = 'utf-8')
    sheet = workbook.add_sheet('export',cell_overwrite_ok=True)

    # 写上字段信息
    for field in range(0,len(fields)):
        sheet.write(0,field,fields[field][0])

    # 获取并写入数据段信息
    row = 1
    col = 0
    for row in range(1,len(results)+1):
        for col in range(0,len(fields)):
            if (results[row-1][3] == results[row-2][3] ):
                if ( col == len(fields)-1 or col == len(fields)-3 or col == len(fields)-4 ):
                    print "\t",
                    pass
                else:
                    sheet.write(row,col,u'%s'%results[row-1][col])
                    print u'%s'%results[row-1][col],
            else:
                sheet.write(row,col,u'%s'%results[row-1][col])
                print u'%s'%results[row-1][col],
        print ''

    export_file_name="export_%s.xls" % datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    print "please check %s" % export_file_name
    workbook.save(export_file_name)
