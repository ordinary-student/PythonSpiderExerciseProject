import requests
import json
from datetime import datetime
from openpyxl import Workbook


# 获取视频信息
def get_video_message(avid):

    # API地址
    url = "https://api.bilibili.com/x/web-interface/view?aid=" + str(avid)

    # 请求头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3741.400 QQBrowser/10.5.3863.400'
    }

    # 获取结果
    result = requests.get(url=url, headers=headers).text
    # 转为json
    data = json.loads(result)
    # 状态码
    code = int(data['code'])

    # 判断
    if code != 0:  # 获取视频信息失败
        return [avid, code, '', '', '', '', '', '', '', '',
                '', '', '', '', '', '', '', '', '']
    else:
        # bv号
        bvid = data['data']['bvid']
        # 标题
        title = data['data']['title']
        # 简介
        desc = data['data']['desc']
        # 封面
        pic = data['data']['pic']
        # 作者id
        uid = data['data']['owner']['mid']
        # 作者名称
        name = data['data']['owner']['name']
        # 作者头像
        face = data['data']['owner']['face']
        # 播放量
        view = data['data']['stat']['view']
        # 弹幕数
        danmu = data['data']['stat']['danmaku']
        # 评论数
        reply = data['data']['stat']['reply']
        # 点赞数
        like = data['data']['stat']['like']
        # 点踩数
        dislike = data['data']['stat']['dislike']
        # 投币数
        coin = data['data']['stat']['coin']
        # 收藏数
        favorite = data['data']['stat']['favorite']
        # 分享数
        share = data['data']['stat']['share']
        # 上传日期
        pubdate = datetime.fromtimestamp(int(data['data']['pubdate']))
        # 地址
        url = 'https://www.bilibili.com/video/'+bvid

        # 视频信息
        video = [avid, code, bvid, title, desc, pic, uid, name, face, view,
                 danmu, reply, like, dislike, coin, favorite, share, pubdate, url]
        # 返回
        return video


# 写入EXCEL文件
def write(video_list):
    # 新建工作簿文件
    wb = Workbook()
    # 新建表格
    sheet = wb.create_sheet('bilibili视频考古', index=0)
    # 写入表格头
    head = ['av号', '状态码', 'bv号', '标题', '简介', '封面', '作者id', '作者名称', '作者头像',
            '播放量', '弹幕数', '评论数', '点赞数', '点踩数', '硬币数', '收藏数', '分享数', '上传日期', '视频地址']
    sheet.append(head)
    # 遍历
    for video in video_list:
        # 添加一行
        sheet.append(video)
    # 保存文件
    wb.save('bilibili视频考古.xlsx')


# 考古
def dig(x, y):
    print('bilibili视频考古开始！')
    # 视频信息列表
    video_list = []
    # 按顺序考古
    for avid in range(x, y+1):
        print('正在考古av'+str(avid)+'...')
        # 获取视频信息
        video = get_video_message(avid)
        # 加入列表
        video_list.append(video)
    # 写入文件
    write(video_list)
    print('bilibili视频考古完成！')


# 主入口
if __name__ == "__main__":
    # 从av1考古到av200
    dig(1, 200)
