from lxml import etree
from xlutils.copy import copy
import requests as req
import xlrd
import xlwt
import sys
import re
import os

head = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/83.0.4103.116 Safari/537.36 Edg/83.0.478.56'}  # 构造请求头

PARAMS = {'header': head}

html = req.get('http://zfcg.nanning.gov.cn//sjcggg/index.htm', params=PARAMS)  # 首次请求

selector = etree.HTML(html.text)

total_page = selector.xpath('//select/option/text()')  # 获取页数


def page_n(page):

    url = 'http://zfcg.nanning.gov.cn//sjcggg/index'
    if page == 1:
        url = url + '.htm'
    else:
        url = url + '_' + str(page) + '.htm'
    data_source = req.get(url, params=PARAMS)
    entity = etree.HTML(data_source.text)

    links = entity.xpath('//div[@class="f-left"]/a/@href')  # 提取a链接里的公告链接
    titles = entity.xpath('//div[@class="f-left"]/a/@title')
    prep = []
    key_words = ['网络', '计算机', '机房', '服务器', '等级保护']  # 按关键字抓取，可根据需要修改，标点符号注意使用英文符号
    index = 0
    for title in titles:
        for kw in key_words:
            if kw in title:
                prep.append(index)

        index += 1

    s1 = range(1, len(links) + 1)
    s2 = [str(x) for x in s1]

    while True:
        index = 1
        for title in titles:
            print(str(index) + ' ' + title + ' ' + links[index - 1]+'\n')
            index += 1
        print('\n当前第' + str(page) + '页,共' + str(len(total_page)) + '页')
        select = input('输入标题左侧的序号抓取对应的标信息,输入0抓取本页关键字项,n查看下一页,m返回主菜单：')

        if select in s2:
            capture(links[int(select) - 1])

        elif select == '0':
            if len(prep) == 0:
                print('本页找不到关键信息，请尝试按序号下载或查看下一页\n\n')
                os.system('pause')
                os.system('cls')
            else:
                for i in prep:
                    capture(links[int(i)])
        elif select == 'n':
            os.system("cls")
            next_page(page + 1)

        elif select == 'm':
            menu()
        else:
            print('你的输入选项有误,请重新输入')


def next_page(page):
    page_n(page)


def capture(link):
    global times
    data_source = req.get(link, params=PARAMS)
    document = etree.HTML(data_source.text)
    t_position = 5
    excel_obj = []

    try:
        release_date = document.xpath('//div[@class="padding5 TxtCenter top10  Gray"]/text()')
        date = date_detail(release_date[0])  # 发布日期
        # print(date)
        excel_obj.append(date)
        rules = '//p[@class="cjk" and position()=3]/font[@size="2"]/text()'

        project = document.xpath(rules)
        if len(project) == 0:
            rules = '//p[@class="cjk" and position()=3]/a[@name="CgggSHEntity_XMBH_0"]/span[@lang="EN-US"]/text()'
            project = document.xpath(rules)
        # print(project)
        project_name = get_project_name(project[1])  # 项目名称
        excel_obj.append(project_name)

        project_items = document.xpath('//p[@class="cjk" and position()=3]/font[@face="Calibri, sans-serif"]/span['
                                       '@lang="en-US"]/font[ '
                                       '@face="Verdana, sans-serif"]/font[@size="2"]/text()')  # 第三个p标签

        project_number = project_items[0]  # 项目编号
        project_budget = get_budget(project_items)  # 预算
        excel_obj.append(project_number)
        excel_obj.append(project_budget)

        # print(project_number)
        # print(project_budget)

        units = document.xpath('//p[@class="cjk" and position()=1]/font[@size="2"]/text()')  # 第一个p标签
        # print(get_unit(purchase, 1))
        purchase_unit = get_unit(units, 0)  # 采购单位
        proxy_unit = get_unit(units, 1)  # 招标代理机构
        excel_obj.append(purchase_unit)
        excel_obj.append(proxy_unit)

        # print(purchase_unit)
        # print(proxy_unit)

        m_list = document.xpath('//p[@class="cjk" and position()=' + str(t_position) + ']/font[@size="2"]/text()')

        while not '五、投标截止时间：' in m_list and t_position < 20:  # 没有找到关键字，p元素向后移
            t_position = t_position + 1
            m_list = document.xpath('//p[@class="cjk" and position()=' + str(t_position) + ']/font[@size="2"]/text()')
        # print(m_list)

        times = document.xpath('//p[@class="cjk" and position()=' + str(t_position) + ']/font[@face="Calibri, '
                                                                                      'sans-serif"]/span[ '
                                                                                      '@lang="en-US"]/font[ '
                                                                                      '@face="Verdana, sans-serif"]/font['
                                                                                      '@size="2"]/text()')
    except Exception as e:
        print(link + ' 抓取过程出现异常,可能是由于网页结构改动所致，请先尝试手动下载或换一另一个标信息抓取\n\n')
        return

    end_time = get_time(times, 0)  # 投标截止时间
    bid_time = get_time(times, 1)  # 开标时间
    excel_obj.append(end_time)
    excel_obj.append(bid_time)

    # print('结束时间：' + end_time)
    # print('开标时间：' + bid_time)

    place = get_bid_info(m_list, times)  # 开标地点
    excel_obj.append(place)
    # print('开标地点：' + place)
    excel_obj.append(link)  # 链接

    save_to_excel(excel_obj)
    print(project_number + ' ' + project_name + '抓取完成' + '\n\n')
    os.system('pause')

def date_detail(date_str):  # 提取日期
    tmp = re.findall(r"期(.+)查", date_str)
    d = tmp[0].split(' ')
    return d[0][1:len(d[0])]


def get_project_name(project):  # 提取项目名称
    p_str = project.strip().split('：')  # 中文符号
    return p_str[1]


def get_budget(items):  # 项目预算金额匹配
    for item in items:
        if re.match(r'^([1-9]\d{0,9}|0)([.]?|(\.\d{1,2})?)$', item):
            return item


def get_unit(items, option):
    groups = items[0].split('委')
    if option == 0:  # 获取采购单位
        return groups[0][1:len(groups[0])]
    else:  # 获取代理机构
        tmp = groups[1].split('拟')
        return tmp[0][2:len(tmp[0])]


def get_time(_times, option):
    time = ''
    if option == 0:
        time = _times[0] + "年" + _times[1] + "月" + _times[2] + "日" + _times[3] + "时" + _times[4] + "分"
    else:
        time = _times[8] + "年" + _times[9] + "月" + _times[10] + "日" + _times[11] + "时" + _times[12] + "分"

    return time


def get_bid_info(bids, times):
    tmp = ''
    index = 0
    for info in bids:
        if '开标地点' in info:
            tmp = info
            break
        index += 1

    index += 1

    pls = tmp.split('：')
    bid_place = pls[1]
    bid_place += times[13]
    bid_place += bids[index]
    index += 1
    bid_place += times[14]
    bid_place += bids[index]
    index += 1
    bid_place += times[15]
    bid_place += bids[index]

    return bid_place


def save_to_excel(obj):
    excel = 'result.xls'
    if os.path.exists(excel):
        tmp_excel = xlrd.open_workbook(excel)
        rows = tmp_excel.sheets()[0].nrows
        excel_file = copy(tmp_excel)
        next_row = rows
        table = excel_file.get_sheet(0)
        col_index = 0
        for col_obj in obj:
            table.write(next_row, col_index, col_obj)
            col_index += 1

        excel_file.save(excel)

    else:
        col_names = ['发布日期', '项目名称', '项目编号', '项目预算', '采购单位', '招标代理机构', '投标截止时间', '开标时间', '开标地点', '链接']
        col_index = 0
        book = xlwt.Workbook()
        sheet = book.add_sheet('Sheet1')  # 添加工作页

        for name in col_names:
            sheet.write(0, col_index, name)
            sheet.write(1, col_index, obj[col_index])
            col_index += 1
        book.save(filename_or_stream=excel)


def menu():
    while True:
        print('\t\t********************欢迎使用南宁市政府集中采购中心自动爬虫脚本********************')
        print('\t\t*             Powered By @Zhengzuo cephmoac@gmail.com                     *')
        print('\t\t*             Data source capture from http://zfcg.nanning.gov.cn/        *')
        print('\t\t*                1.开始使用(从第一页开始)                                     *')
        print('\t\t*                2.自定义开始页                                              *')
        print('\t\t*                3.退出脚本                                                 *')
        print('\t\t****************************************************************************')
        select = input('\n\t\t请选择以上带有数字的选项，按[Enter]执行：')
        if int(select) == 1:
            page_n(1)
        elif int(select) == 2:
            select = input('\n\t\t请输入页数，按[Enter]执行：')
            page_n(select)
        elif int(select) == 3:
            sys.exit(0)
        else:
            print('你输入的选项有误')


menu()
