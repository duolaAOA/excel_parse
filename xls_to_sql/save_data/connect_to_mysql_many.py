# -*-coding:utf-8 -*-
# time: 2018.3.22

import pymysql
import datetime
import logging


class get_Mysql(object):
    def __init__(self, dbname, xls_name):
        self.dbname = dbname
        self.T = datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d%H%M")
        # 数据表的名称
        self.table_name = '{}'.format(xls_name)
        # 连接接数据库
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

    def create_table(self):
        """
        创建数据表的函数，表格名称按照时间和关键词命名
        """
        sql = '''CREATE TABLE `{tbname}` (
              `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
              {biddingMethod} varchar(20) DEFAULT NULL COMMENT '招标方式',
              {biddingCycle} varchar(20) DEFAULT NULL COMMENT '招标周期',
              {identifier} varchar(255) DEFAULT NULL COMMENT '编号',
              {hbrainCategory} varchar(255) DEFAULT NULL COMMENT '物料总类',
              {orderedItem} varchar(255) DEFAULT NULL COMMENT '物料名称',
              {orderedItemIdentifier} varchar(255) DEFAULT NULL COMMENT '物料编码',
              {itemCondition} varchar(255) DEFAULT NULL COMMENT '规格',
              {customer} varchar(255) CHARACTER SET utf8 COLLATE utf8_unicode_ci DEFAULT NULL COMMENT '使用企业',
              {seller} varchar(255) DEFAULT NULL COMMENT '中标人／供应商',
              {orderDate} datetime DEFAULT NULL COMMENT '开标时间／单据日期',
              {totalPrice} varchar(255) DEFAULT NULL COMMENT '中标金额／采购总价',
              PRIMARY KEY (`id`)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8;
'''
        try:
            self.cursor.execute(sql.format(tbname=self.table_name, biddingMethod="biddingMethod",
                                           biddingCycle="biddingCycle", identifier="identifier",
                                           hbrainCategory="hbrainCategory", orderedItem="orderedItem",
                                           orderedItemIdentifier="orderedItemIdentifier", itemCondition="itemCondition",
                                           customer="customer", seller="seller", orderDate="orderDate",
                                           totalPrice="totalPrice",))
        except Exception as e:
            logging.warning(e)
            print("创建数据表失败，表格可能已经存在！", e)
        else:
            self.conn.commit()
            print("成功创建一个数据表，名称是{}".format(self.table_name))

    def insert_data(self, data):
        """数据插入"""
        insert_sql = '''INSERT INTO `{tbname}`(biddingMethod, biddingCycle, identifier, hbrainCategory, orderedItem,
                                                orderedItemIdentifier, itemCondition, customer, seller, orderDate,
                                                totalPrice) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)                                  
        '''
        try:
            self.cursor.executemany(insert_sql.format(tbname=self.table_name), data)

        except Exception as e:
            self.conn.rollback()
            print("Insert data failure, cause：", e)
            logging.warning(e)

        else:
            self.conn.commit()
            print('Insert a data successfully!')

    def close_table(self):
        print("End of data insert !")
        self.cursor.close()
        self.conn.close()


if __name__ == '__main__':
    """
    测试
    """
    data = {'biddingMethod': '线上招标', 'biddingCycle': '月度招标', 'identifier': 'ZB-ZY-ZY-1802002-1',
            'hbrainCategory': '中药', 'orderedItem': '月度招标', 'orderedItemIdentifier': 'ZB-ZY-ZY-1802002-1',
            'itemCondition': '月度招标', 'customer': '中药二厂', 'seller': '安国义通中药材有限公司',
            'orderDate': '2018/2/6', 'price': '21', 'totalUnit': '1000', 'totalPrice': '0.53'}

    my = get_Mysql('market_price',
                   'test1')
    my.create_table()
    my.insert_data(data)
    my.close_table()

