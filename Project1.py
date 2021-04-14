# --*--coding:utf-8 --*--
# Author: Mu Runlin

import json
import urllib.request, urllib.error  # 制定URL，获取网页数据
import xlwt


def main():
    baseurl = "http://datainterface.eastmoney.com/EM_DataCenter/JS.aspx?cb=datatable3791565&type=GJZB&sty=ZGZB&js=(%7Bdata%3A%5B(x)%5D%2Cpages%3A(pc)%7D)&p=2&ps=20&mkt=19&pageNo=2&pageNum=2&_=1618382432294"
    # 1.爬取网页
    datalist = getData(baseurl)
    print(datalist)

    savePath = ".\\CPI.xls"
    # 2.解析数据

    # 3.保存数据
    saveData(datalist, savePath)


# 1.
# 2.边爬取边解析。理论上：逐一解析数据
def getData(baseurl):
    html = askURL(baseurl)
    # 2.解析数据
    cpi_data = str(html).lstrip(r"datatable3791565({data:[ ").rstrip(r"],pages:8})")  # 转为字符串类型，去除非json格式数据(去头去尾)
    print(type(cpi_data))
    print("m" + cpi_data)

    # 将字符串转换为列表
    cpi = cpi_data.replace("\"", '')
    print(cpi)
    cpi2 = cpi.split(',')
    print(cpi2)
    print(type(cpi2))

    return cpi2


# 得到某一URL的网页内容
def askURL(url):
    head = {  # 用户代理，伪装成一个浏览器，模拟浏览器头部信息，向目标服务器发送消息，以及告知自己的接收水平
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36 Edg/89.0.774.75"
    }

    request = urllib.request.Request(url, headers=head)  # 封装request头部信息
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e, "code")
        if hasattr(e, "reason"):
            print(e, "reason")
    return html


# 3.
def saveData(datalist, savePath):
    print("save...")

    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('CPI指数', cell_overwrite_ok=True)
    col = ("月份", "当月", "同比增长", "环比增长", "累计", "当月", "同比增长", "环比增长", "累计", "当月", "同比增长", "环比增长", "累计")

    print(datalist[2])

    for j in range(0, 13):
        sheet.write(0, j, col[j])     # 生成表格头

    print(len(datalist))
    for i in range(0, len(datalist)):
        print("第%d个" % (i + 1))
        data = datalist[i]
        sheet.write(int((i%13)+1), int(i/13), data)   # 写入excel

    book.save(savePath)


if __name__ == "__main__":
    main()
    print('finish')

