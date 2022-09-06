import requests
import re
import xlrd
import xlwt

def get_Htmltext(url):
    """根据url，获得HTML文本"""
    try:
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36"
        header = {
            "user-agent": user_agent
        }
        resp = requests.get(url, headers=header)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding
        return resp.text
    except:
        return ""

def expression():
    """定义正则表达式"""
    obj = re.compile(r'"f2":(?P<f2>.*?),"f3":(?P<f3>.*?),'
                     r'"f4":(?P<f4>.*?),'
                     r'"f5":(?P<f5>.*?),'
                     r'"f6":(?P<f6>.*?),'
                     r'"f7":(?P<f7>.*?),'
                     r'"f8":(?P<f8>.*?),'
                     r'"f9":(?P<f9>.*?),'
                     r'"f10":(?P<f10>.*?),'
                     r'"f11":(?P<f11>.*?),'
                     r'"f12":(?P<f12>.*?),'
                     r'"f13":(?P<f13>.*?),'
                     r'"f14":(?P<f14>.*?),'
                     r'"f15":(?P<f15>.*?),'
                     r'"f16":(?P<f16>.*?),'
                     r'"f17":(?P<f17>.*?),'
                     r'"f18":(?P<f18>.*?),'
                     r'"f20":(?P<f20>.*?),'
                     r'"f21":(?P<f21>.*?),'
                     r'"f22":(?P<f22>.*?),'
                     r'"f23":(?P<f23>.*?),'
                     r'"f24":(?P<f24>.*?),'
                     r'"f25":(?P<f25>.*?),'
                     , re.S)
    objn = re.compile(r'{"f1":.*?}')
    return obj,objn

def perse_htmltext(page_content,total_diclenlist,obj,objn):
    """解析数据，提取文本"""
    #使用正则提取文本
    result = obj.finditer(page_content)
    nums = objn.findall(page_content)
    numsdic_len = len(nums)
    total_diclenlist.append(numsdic_len)
    sum_total = sum(total_diclenlist)
    return result,sum_total

def creat_sheet():
    """创建一个模板Excel文件，并将其信息放到新表"""
    #选择模板表
    book = xlrd.open_workbook("科创版.xls")
    sheet = book.sheet_by_index(0)
    # 表的行数和列数
    st_r, st_c = sheet.nrows, sheet.ncols
    nwb = xlwt.Workbook(encoding="utf-8")
    nws = nwb.add_sheet("创业版股票数据")
    for r in range(st_r):
        for c in range(st_c):
            nws.write(r, c, sheet.cell(r, c).value)
    return nwb,nws

def save_excel(result,nws,sum_total,h):
    """将信息提取到Excel文件"""
    for it in result:
        dic = it.groupdict()
        if h < (sum_total + 1):
            nws.write(h, 0, eval(dic["f12"]))
            nws.write(h, 1, eval(dic["f14"]))
            nws.write(h, 2, dic["f2"])
            nws.write(h, 3, dic["f3"])
            nws.write(h, 4, dic["f4"])
            nws.write(h, 5, dic["f5"])
            nws.write(h, 6, dic["f6"])
            nws.write(h, 7, dic["f7"])
            nws.write(h, 8, dic["f8"])
            nws.write(h, 9, dic["f11"])
            nws.write(h, 10, dic["f15"])
            nws.write(h, 11, dic["f16"])
            nws.write(h, 12, dic["f17"])
            nws.write(h, 13, dic["f18"])
            nws.write(h, 14, dic["f20"])
            nws.write(h, 15, dic["f21"])
            nws.write(h, 16, dic["f22"])
            nws.write(h, 17, dic["f9"])
            nws.write(h, 18, dic["f23"])
            nws.write(h, 19, dic["f4"])
            nws.write(h, 20, dic["f25"])
            h += 1
    return h


def main():
    """定义一个主程序"""
    #定义初始变量
    #parsecontent=input("请选择要爬取的板块：")
    total_page = 70
    h = 1
    total_diclenlist = []
    # 创建一个模板Excel文件，并将其信息放到新表
    nwb, nws = creat_sheet()
    #正则表达
    obj,objn = expression()
    for p in range(1, total_page + 1):
        page = p
        """获取HTML文本"""
        """板块"""
        url = f"http://30.push2.eastmoney.com/api/qt/clist/get?cb=jQuery112408850200122749612_1631165047301&pn={page}&pz=20&po=1&np=1&ut=bd1d9ddb04089700cf9c27f6f7426281&fltt=2&invt=2&fid=f3&fs=m:0+t:80&fields=f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f12,f13,f14,f15,f16,f17,f18,f20,f21,f23,f24,f25,f22,f11,f62,f128,f136,f115,f152"
        # 提取HTML文本
        page_content=get_Htmltext(url)
        # 解析到需要的HTML文本信息
        result,sum_total = perse_htmltext(page_content, total_diclenlist,obj,objn)
        # 将信息提取到Excel文件
        h=save_excel(result, nws, sum_total,h)
    nwb.save("创业版数据表.xls")
    print("保存成功")

main()







