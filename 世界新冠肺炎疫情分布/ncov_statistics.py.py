import requests
import json
import jsonpath
from openpyxl import Workbook


# 获取数据
def get_data():

    # API地址
    url = 'https://api.inews.qq.com/newsqa/v1/automation/foreign/country/ranklist'

    # 添加请求头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3741.400 QQBrowser/10.5.3863.400'
    }

    # 获取结果
    res = requests.get(url=url,  headers=headers).text
    # 转为json
    data = json.loads(res)

    # 获取数据
    # 国名
    name = jsonpath.jsonpath(data, '$..name')
    # 总确诊人数
    confirm = jsonpath.jsonpath(data, '$..confirm')
    # 新增确诊
    confirmAdd = jsonpath.jsonpath(data, '$..confirmAdd')
    # 现存确诊人数
    nowConfirm = jsonpath.jsonpath(data, '$..nowConfirm')
    # 治愈
    heal = jsonpath.jsonpath(data, '$..heal')
    # 死亡人数
    dead = jsonpath.jsonpath(data, '$..dead')
    # 确诊对比
    confirmCompare = jsonpath.jsonpath(data, '$..confirmCompare')
    # 现存确诊对比
    nowConfirmCompare = jsonpath.jsonpath(data, '$..nowConfirmCompare')
    # 治愈对比
    healCompare = jsonpath.jsonpath(data, '$..healCompare')
    # 死亡对比
    deadCompare = jsonpath.jsonpath(data, '$..deadCompare')

    # 配对
    data_list = zip(name, confirm, confirmAdd, nowConfirm,
                    heal, dead, confirmCompare, nowConfirmCompare, healCompare, deadCompare)
    # 转列表
    ll = list(data_list)
    # 总数据列表
    result_list = []
    # 遍历
    for l in ll:
        # 元祖转列表
        d = list(l)
        # 加入总列表
        result_list.append(d)
        print(d)

    # 返回
    return result_list


# 写入EXCEL文件
def write(data_list):

    # 新建工作簿文件
    wb = Workbook()
    # 新建表格
    sheet = wb.create_sheet('海外疫情状况', index=0)
    # 添加表头
    table_head = ['国家', '总确诊', '新增确诊', '现存确诊', '治愈', '死亡',
                  '确诊对比', '现存确诊对比', '治愈对比'	, '死亡对比', '治愈率', '死亡率']
    sheet.append(table_head)
    # 遍历
    for data in data_list:
        # 添加一行
        sheet.append(data)
    # 保存文件
    wb.save('世界新冠肺炎疫情状况统计.xlsx')
    print('写入文件完成')


# 主函数
if __name__ == "__main__":
    # 获取数据
    data_list = get_data()
    # 写入文件
    write(data_list)
