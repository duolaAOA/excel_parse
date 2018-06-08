# -*-coding:utf-8 -*-
# time: 2018.1.23
# python2.7.14

import os
import sys
import xlrd

from save_data.connect_to_mysql import get_Mysql

reload(sys)
sys.setdefaultencoding('utf-8')


class xlsx_to_Mysql(object):
    # 文件存放地址
    def __init__(self, dbname, file_path):
        self.dbname = dbname    # 数据库名
        self.filepath = file_path
        self.data = {}  # 解析后的字段以字典存放

    def find_file_path(self):
        """
        文件地址
        """
        # 将所需要批处理的文件放到指定路径下
        # 循环处理每个文件
        for dirname, dirnames, filenames in os.walk(self.filepath):
            for filename in filenames:
                file = os.path.join(dirname, filename)
                xls_name = (filename.split('.')[0]).decode('gbk')
                self.mysql = get_Mysql(self.dbname, xls_name)
                self.mysql.create_table()  # 建表
                self.read_xls(file)
        self.mysql.close_table()

    def read_xls(self, file):
        # 获取xls文件数据
        parse_xls = xlrd.open_workbook(file)
        first_sheet = parse_xls.sheet_by_index(0)
        for each_row in range(1, first_sheet.nrows):
            # _id = first_sheet.cell(each_row, 0).value
            headline = first_sheet.cell(each_row, 1).value
            orgpermitted_name = first_sheet.cell(each_row, 2).value
            # 日期格式处理
            datePublished = xlrd.xldate_as_datetime(first_sheet.cell(each_row, 3).value, 0)
            copyrightHolder = first_sheet.cell(each_row, 4).value
            referenceUrl = first_sheet.cell(each_row, 5).value
            # self.data['_id'] = _id
            self.data['headline'] = headline
            self.data['orgpermitted_name'] = orgpermitted_name
            self.data['datePublished'] = datePublished
            self.data['copyrightHolder'] = copyrightHolder
            self.data['referenceUrl'] = referenceUrl
            self.mysql.insert_data(self.data)


if __name__ == '__main__':
    file_path = './data'         # 存放数据的路径
    db_name = 'market_price'    # 数据库名
    xls = xlsx_to_Mysql(db_name, file_path)
    xls.find_file_path()

