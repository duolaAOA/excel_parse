# -*-coding:utf-8 -*-

import json
import os
import re



def remove_dirty_data(line):
    """移除数据中非json格式字符"""
    remove_left = re.sub('ObjectId\(', '', line)
    remove_right = re.sub('\)', '', remove_left)
    return remove_right


def convert(filename):
    """字段提取"""
    print("Opening TXT file: ", os.path.basename(filename))
    with open(filename, 'rb') as f:
        for line in f:
            if line == '\n':
                pass
            else:
                new_line = remove_dirty_data(line.decode('utf-8'))
                hash_value = new_line.strip()
                try:
                    json_str = json.loads(hash_value)
                    # 产品id
                    _id = json_str["_id"]
                    # 产品版本
                    version = json_str["meta_version"]
                    # 产品类型
                    product_type = json_str["download_data"]["parsed_data"]["product_type"]
                    # 产品供应商
                    producer = json_str["download_data"]["parsed_data"]["producer"]
                    # 产品价格
                    price_unit = json_str["download_data"]["parsed_data"]["price"]
                    # 价格不带单位 匹配：2100，02100，02100.0，2100.0
                    price = re.match('[1-9]\d*\.\d*|0\.\d*[1-9]\d*$|(\d*)(.*?)', price_unit).group()
                    # 价格单位
                    unit = re.split('[1-9]\d*\.\d*|0\.\d*[1-9]\d*$|(\d*)(.*?)', price_unit)[-1]
                    # 产品推送日期
                    datePublished = json_str["download_data"]["parsed_data"]["datePublished"]
                    # 来源地
                    fromLocation = json_str["download_data"]["parsed_data"]["fromLocation"]
                    # 产品级别
                    itemCondition = json_str["download_data"]["parsed_data"]["itemCondition"]
                    # 产品价格
                    priceType = json_str["download_data"]["parsed_data"]["priceType"]
                    # 来源链接
                    url = json_str["download_config"]["url"]
                    # 获取方式
                    method = json_str["download_config"]["method"]
                    # 更新时间
                    meta_updated = re.sub('T', ' ', json_str["meta_updated"])
                    print(meta_updated)
                except ValueError as e:
                    raise e


def find_file_path(file_path):
    """
    文件地址
    """
    # 将所需要批处理的文件放到指定路径下
    # 循环处理每个文件
    for dirname, dirnames, filenames in os.walk(file_path):
        for filename in filenames:
            file = os.path.join(dirname, filename)
            convert(file)


if __name__ == '__main__':
    # 存放清洗数据的路径
    file_path = './txt'
    find_file_path(file_path)
