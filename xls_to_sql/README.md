#### 批处理xls文件，字段提取导入mysql
* data文件夹用于存放需要处理的数据
* save_data 为数据库操作， 需要手动更改数据库账号密码
* 在excel_to_mysql 中 函数入口输入数据库名和xls文件存放地址
* table_name 数据表名,  默认为None表示按文件名分别生成数据表,否则为自定义的数据表，并将左右数据保存至一张表

##
* excel_to_mysql_many   修改为Mysql批量插入
* save_data 中  在excel_to_mysql   中接收的data为可迭代的list对象， 元素未tuple，  执行批量插入操作 (executemany)