# -*- coding: utf-8 -*-
import json
import logging
import pymongo
import xlrd


class TaskMake(object):
    """
    任务参数入库(mongo)
    """
    def __init__(self):
        pass

    def __call__(self, *args, **kwargs):
        return self.make_task()

    def mongo_connect(self):
        """
        连接mongo
        :return: table
        """
        mongo_uri = 'mongodb://192.168.0.202:27027'
        # mongo_db = "tender_test"
        mongo_db = "tender_task"
        client = pymongo.MongoClient(mongo_uri)
        db = client[mongo_db]
        # table = db[f"params"]
        table = db[f"task_jie"]
        return table


    def make_task(self):
        """
        任务参数入库存储
        :return:
        """
        for param,i in self.operate_xlrd():
            try:
                table = self.mongo_connect()
                request_dict = eval(param)
                if 'data' in request_dict['list'] and type(request_dict['list']['data']) is dict:
                    request_dict['list']['data'] = json.dumps(request_dict['list']['data'])
                    request_dict['detail']['data'] = json.dumps(request_dict['list']['data'])
                table.insert_one(request_dict)
                logging.warning(f"===插入成功!,行号：{i+1}===")
            except Exception as e:
                logging.warning(f"===当前插入数据有误{e},行号：{i+1}，===")

    def operate_xlrd(self):
        """
        操作读取excel
        :return:
        """
        data = xlrd.open_workbook(r'D:\Desktop\新的表.xlsx', encoding_override='utf-8')
        table = data.sheets()[0] #选定表
        nrows = table.nrows #获取行号
        for i in range(1, nrows):
            alldata = table.row_values(i)  # 输出所有数据
            params = alldata[0]  # 取出表中第一列数据

            yield params,i


if __name__ == '__main__':
    a = TaskMake()
    a()
