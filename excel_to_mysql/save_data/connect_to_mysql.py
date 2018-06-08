# -*-coding:utf-8 -*-

import pymysql
import datetime


class get_Mysql(object):
    def __init__(self, dbname, xls_name):
        self.dbname = dbname
        self.T = datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d%H%M")
        # 数据库表格的名称
        self.table_name = '{}'.format(xls_name)
        # 链接数据库
        self.conn = pymysql.connect(
            host='127.0.0.1',
            user='root',
            password='1219960386',
            port=3306,
            db=self.dbname,
            charset='utf8'
        )
        # 获取游标
        self.cursor = self.conn.cursor()
    # 创建表格的函数，表格名称按照时间和关键词命名

    def create_table(self):
        sql = '''CREATE TABLE `{tbname}` (
          {headline} varchar(255) DEFAULT NULL COMMENT '标题',
          {orgpermitted} varchar(255) DEFAULT NULL COMMENT '企业名称',
          {datePublished} datetime DEFAULT NULL COMMENT '发布时间',
          {copyrightHolder} varchar(255) DEFAULT NULL COMMENT '数据来源',
          {referenceUrl} varchar(255) DEFAULT NULL COMMENT '数据源链接'
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8;
'''
        try:
            self.cursor.execute(sql.format(tbname=self.table_name, headline="headline", orgpermitted="orgpermitted_name", datePublished="datePublished",
                                           copyrightHolder="copyrightHolder", referenceUrl="referenceUrl"))
        except Exception as e:
            print("创建表格失败，表格可能已经存在！", e)
        else:
            self.conn.commit()
            print("成功创建一个表格，名称是{}".format(self.table_name))

    def insert_data(self, data):
        """数据插入"""
        insert_sql = '''INSERT INTO `{tbname}`(headline, orgpermitted_name, datePublished, copyrightHolder, referenceUrl) 
                        VALUES (%s, %s, %s, %s, %s)                                  
        '''
        try:
            self.cursor.execute(insert_sql.format(tbname=self.table_name), (data['headline'], data['orgpermitted_name'],
                                                                            data['datePublished'], data['copyrightHolder'], data['referenceUrl']))

        except Exception as e:
            self.conn.rollback()
            print("Insert data failure, cause：",e)

        else:
            self.conn.commit()
            print('Insert a data successfully!')

    def close_table(self):
        print("End of data insert !")
        self.cursor.close()
        self.conn.close()


if __name__ == '__main__':
    data = {'create_time':'09月15日 22:56','mobile_phone':'来自荣耀V8 脱影而出','content':'qwewq','crawl_time':'2017-09-22 19:15:08'}

    my = get_Mysql('market_price',
                   'fgyugyujh')
    my.create_table()
    # my.insert(data)
    # my.close_table()

