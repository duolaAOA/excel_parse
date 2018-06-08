# -*-coding:utf-8 -*-
# time: 2018.3.22

import os
import xlrd
from datetime import datetime
import time
import logging


from save_data.connect_to_mysql import get_Mysql


class xlsx_to_Mysql(object):
    # 文件存放地址
    def __init__(self, dbname, file_path, table_name=None):
        self.dbname = dbname    # 数据库名
        self.filepath = file_path
        self.table_name = table_name   # 数据表名
        self.data = {}  # 解析后的字段以字典存放
        # 日志
        T = datetime.strftime(datetime.now(), "%Y-%m-%d-%H-%M")
        logging.basicConfig(level=logging.NOTSET,
                            filename='./log/{}.log'.format(T),
                            filemode='w',
                            format='%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s')

    def find_file_path(self):
        """
        文件地址
        """
        # 将所需要批处理的文件放到指定路径下
        # 循环处理每个文件
        for dirname, dirnames, filenames in os.walk(self.filepath):
            for filename in filenames:
                file = os.path.join(dirname, filename)
                xls_name = (filename.split('.')[0]).replace(' ', '')
                if self.table_name is None:
                    self.mysql = get_Mysql(self.dbname, xls_name)
                else:
                    self.mysql = get_Mysql(self.dbname, self.table_name)
                self.mysql.create_table()  # 建表
                self.read_xls(file)
        self.mysql.close_table()

    def read_xls(self, file):
        # 获取xls文件与
        parse_xls = xlrd.open_workbook(file)
        logging.info("当前处理的excel文件名为---{}----".format(file))
        sheet_names = parse_xls.sheet_names()       # 获取所有sheet表名
        for sheet in sheet_names:
            """
            过滤不需要的sheet
            """
            if '登记' in sheet:
                pass
            elif '异常' in sheet:
                pass
            elif 'Sheet' in sheet:
                pass
            elif '2015' in sheet:
                pass
            else:
                parse_sheet = parse_xls.sheet_by_name(sheet)
                self.parse_sheet_content(parse_sheet, sheet)

    def parse_sheet_content(self, parse_sheet, sheet):
        """
        处理每一个sheet
        """
        choice = ['设备', '能源', ]
        logging.info("当前处理的sheet表名为---{}----".format(sheet))
        if sheet in choice:
            # 判断另一种格式
            for each_row in range(5, parse_sheet.nrows):
                biddingMethod = parse_sheet.cell(each_row, 11).value
                biddingCycle = parse_sheet.cell(each_row, 21).value
                identifier = parse_sheet.cell(each_row, 6).value
                hbrainCategory = parse_sheet.cell(each_row, 2).value
                orderedItem = parse_sheet.cell(each_row, 9).value
                orderedItemIdentifier = parse_sheet.cell(each_row, 8).value
                itemCondition = parse_sheet.cell(each_row, 10).value
                customer = parse_sheet.cell(each_row, 13).value
                seller = parse_sheet.cell(each_row, 27).value
                time_isinstance = parse_sheet.cell(each_row, 19).value

                if isinstance(time_isinstance, float):
                    """
                    默认时间为1990-01-09 00:00:00
                    """
                    orderDate = xlrd.xldate_as_datetime(time_isinstance, 0)
                else:
                    t = time.strptime("1990-11-11", "%Y-%m-%d")
                    y, m, d = t[0:3]
                    orderDate = datetime(y, m, d)
                price = parse_sheet.cell(each_row, 29).value
                totalUnit = parse_sheet.cell(each_row, 30).value
                totalPrice = parse_sheet.cell(each_row, 28).value
                self.data['biddingMethod'] = biddingMethod
                self.data['biddingCycle'] = biddingCycle
                self.data['identifier'] = identifier
                self.data['hbrainCategory'] = hbrainCategory
                self.data['orderedItem'] = orderedItem
                self.data['orderedItemIdentifier'] = orderedItemIdentifier
                self.data['itemCondition'] = itemCondition
                self.data['customer'] = customer
                self.data['seller'] = seller
                self.data['orderDate'] = orderDate
                self.data['price'] = price
                self.data['totalUnit'] = totalUnit
                self.data['totalPrice'] = totalPrice
                logging.info(self.data)
                self.mysql.insert_data(self.data)

                # 使用executemany  批量插入Mysql
                # 数据存放类型定义为元组或列表，  加上一个计时器， 多次读取一次插入加快效率
                # data_list = (biddingMethod, biddingCycle, identifier, hbrainCategory, orderedItem,
                #           orderedItemIdentifier, itemCondition, customer, seller, orderDate,
                #           totalPrice, )
                # self.data.append(data_list)
                # self.mysql.insert_data(self.data)
                # self.data.clear()
    else:
            # 有效数据从第5行开始
            for each_row in range(5, parse_sheet.nrows):
                biddingMethod = parse_sheet.cell(each_row, 11).value
                biddingCycle = parse_sheet.cell(each_row, 22).value
                identifier = parse_sheet.cell(each_row, 6).value
                hbrainCategory = parse_sheet.cell(each_row, 2).value
                orderedItem = parse_sheet.cell(each_row, 9).value
                orderedItemIdentifier = parse_sheet.cell(each_row, 8).value
                itemCondition = parse_sheet.cell(each_row, 10).value
                customer = parse_sheet.cell(each_row, 13).value
                seller = parse_sheet.cell(each_row, 28).value
                time_isinstance = parse_sheet.cell(each_row, 20).value

                if isinstance(time_isinstance, float):
                    """
                    默认时间为1900-01-09 00:00:00
                    """
                    orderDate = xlrd.xldate_as_datetime(time_isinstance, 0)
                else:
                    t = time.strptime("1990-11-11", "%Y-%m-%d")
                    y, m, d = t[0:3]
                    orderDate = datetime(y, m, d)
                price = parse_sheet.cell(each_row, 30).value
                totalUnit = parse_sheet.cell(each_row, 31).value
                totalPrice = parse_sheet.cell(each_row, 29).value
                self.data['biddingMethod'] = biddingMethod
                self.data['biddingCycle'] = biddingCycle
                self.data['identifier'] = identifier
                self.data['hbrainCategory'] = hbrainCategory
                self.data['orderedItem'] = orderedItem
                self.data['orderedItemIdentifier'] = orderedItemIdentifier
                self.data['itemCondition'] = itemCondition
                self.data['customer'] = customer
                self.data['seller'] = seller
                self.data['orderDate'] = orderDate
                self.data['price'] = price
                self.data['totalUnit'] = totalUnit
                self.data['totalPrice'] = totalPrice
                logging.info(self.data)
                self.mysql.insert_data(self.data)


if __name__ == '__main__':
    file_path = './data'         # 存放数据的路径
    db_name = 'market_price'    # 数据库名

    xls = xlsx_to_Mysql(db_name, file_path)
    xls.find_file_path()
