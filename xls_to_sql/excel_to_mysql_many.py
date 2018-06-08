# -*-coding:utf-8 -*-
# time: 2018.3.22

import os
import xlrd
from datetime import datetime
import time
import logging
import re


from save_data.connect_to_mysql_many import get_Mysql


class xlsx_to_Mysql(object):
    # 文件存放地址
    def __init__(self, dbname, file_path, table_name=None):
        self.dbname = dbname    # 数据库名
        self.filepath = file_path
        self.table_name = table_name   # 数据表名
        self.data = []  # 解析后的字段以列表存放

        self.insert_count = 0  # 计数器  每一次插入的条数
        self.max_count = 2000   # 列表每次存放的最大数据量

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
        # 获取xls文件
        try:
            global default_year
            default_year = re.search('\d{4}', file).group()
        except AttributeError as e:
            logging.error(e, "该xls文件名不包含四位的年份数字")
            print(e, "该xls文件名不包含四位的年份数字")

        parse_xls = xlrd.open_workbook(file)
        logging.info("当前处理的excel文件名为---{}----".format(file))
        sheet_names = parse_xls.sheet_names()       # 获取所有sheet表名

        filter_sheet = '登记|异常|Sheet|2015'   
        for sheet in sheet_names:
            if re.search(filter_sheet, sheet):
                pass

            else:
                parse_sheet = parse_xls.sheet_by_name(sheet)
                self.parse_sheet_content(parse_sheet, sheet)

    def parse_sheet_content(self, parse_sheet, sheet):
        """
        处理每一个sheet
        """
        filter_sheet = '设备|能源|杂品|工程|服务'

        logging.info("当前处理的sheet表名为---{}----".format(sheet))

        if re.search(filter_sheet, sheet):
            self._get_field_inert_one(parse_sheet)

        else:
            self._get_field_inert_two(parse_sheet)

    def _get_field_inert_one(self, parse_sheet):
        """
        对第一种格式的处理
        :return:
        """
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

                orderDate = xlrd.xldate_as_datetime(time_isinstance, 0)

            elif re.search('/', str(time_isinstance)):
                orderDate = Deal_Time.time_with_one(str(time_isinstance))

            elif re.search('-', str(time_isinstance)):
                orderDate = Deal_Time.time_with_two(str(time_isinstance))

            elif str(time_isinstance).isdigit():
                orderDate = Deal_Time.time_with_three(str(time_isinstance))

            elif re.search('月', str(time_isinstance)):
                orderDate = Deal_Time.time_with_four(str(time_isinstance))

            elif re.search('.', str(time_isinstance)):
                orderDate = Deal_Time.time_with_five(str(time_isinstance))

            else:
                orderDate = Deal_Time.default_time()

            totalPrice = parse_sheet.cell(each_row, 28).value

            data_list = (biddingMethod, biddingCycle, identifier, hbrainCategory, orderedItem,
                         orderedItemIdentifier, itemCondition, customer, seller, orderDate, totalPrice)

            if self.insert_count < self.max_count:
                print(self.insert_count)
                self.data.append(data_list)
                logging.info(data_list)
                self.insert_count += 1
            else:
                self.mysql.insert_data(self.data)
                logging.info(self.data)
                self.data.clear()
                self.insert_count = 0

    def _get_field_inert_two(self, parse_sheet):
        """
        对第二种格式的处理
        # 有效数据从第5行开始
        """
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
                orderDate = xlrd.xldate_as_datetime(time_isinstance, 0)

            elif re.search('/', str(time_isinstance)):
                orderDate = Deal_Time.time_with_one(str(time_isinstance))

            elif re.search('-', str(time_isinstance)):
                orderDate = Deal_Time.time_with_two(str(time_isinstance))

            elif str(time_isinstance).isdigit():
                orderDate = Deal_Time.time_with_three(str(time_isinstance))

            elif re.search('月', str(time_isinstance)):
                orderDate = Deal_Time.time_with_four(str(time_isinstance))

            elif re.search('.', str(time_isinstance)):
                orderDate = Deal_Time.time_with_five(str(time_isinstance))

            else:
                orderDate = Deal_Time.default_time()

            totalPrice = parse_sheet.cell(each_row, 29).value

            data_list = (biddingMethod, biddingCycle, identifier, hbrainCategory, orderedItem,
                         orderedItemIdentifier, itemCondition, customer, seller, orderDate, totalPrice)

            if self.insert_count < self.max_count:
                print(self.insert_count)
                self.data.append(data_list)
                logging.info(data_list)
                self.insert_count += 1
            else:

                self.mysql.insert_data(self.data)
                logging.info(self.data)
                self.data.clear()
                self.insert_count = 0


class Deal_Time(object):
    """
    日期处理
    """
    def __init__(self):
        self.t = "1990-11-11"
        self.t_layout = "%Y-%m-%d"

    @staticmethod
    def default_time():
        """
        如果表日期字段为空
        :return:  2000-11-11 00:00:00
        """
        return time.strptime(self.t, self.t_layout)

    @staticmethod
    def time_with_one(time_isinstance):
        """
        处理类似  2017/2/15
        """
        return datetime.strptime(time_isinstance, "%Y/%m/%d")

    @staticmethod
    def time_with_two(time_isinstance):
        """
        处理类似  2017-2-15
        """
        return datetime.strptime(time_isinstance, self.t_layout)

    @staticmethod
    def time_with_three(time_isinstance):
        """
        处理类似  2017-2-15
        """
        return datetime.strptime(time_isinstance, "%Y%m%d")

    @staticmethod
    def time_with_four(time_isinstance):
        """
        处理类似  "4月15日"
        """
        tmp_time_isinstance = time_isinstance.split("月")
        month = tmp_time_isinstance[0]
        day = tmp_time_isinstance[-1].replace("日", '')
        new_datetime = "{0}-{1}-{2}".format(default_year, month, day)
        return datetime.strptime(new_datetime, self.t_layout)

    @staticmethod
    def time_with_five(time_isinstance):
        """
        处理类似  2017-2-15
        """
        return datetime.strptime(time_isinstance, "%Y.%m.%d")


if __name__ == '__main__':
    file_path = './data'         # 存放数据的路径
    db_name = 'market_price'    # 数据库名

    xls = xlsx_to_Mysql(db_name, file_path)
    xls.find_file_path()
