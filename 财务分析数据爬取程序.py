import requests
import re
from bs4 import BeautifulSoup
from openpyxl import Workbook

"""编辑url信息"""

def edit_url(month,year,stock_code="002460"):
    month_day = {"3": "31", "6": "30", "9": "30", "12": "31"}
    i = 0
    month_list = []
    year_list = []
    month_list.append(month)
    year_list.append(str(year))
    while i < 4:
        month = month - 3
        if month == 0:
            month = 12
            year = year - 1
        year_list.append(str(year))
        month_list.append(month)
        i += 1
    if time is True:
        url_zfczb = f"http://emweb.eastmoney.com/PC_HSF10/NewFinanceAnalysis/zcfzbAjaxNew?companyType=4&reportDateType=1&" \
                    f"reportType=1&dates={time}-12-31%2C{time-1}-12-31%2C{time-2}-12-31%2C{time-3}-12-31%2C{time-4}-12-31" \
                    f"&code={stock_demo1}"
    elif year is True:
        # 按报告期查询
        url_zfczb=f"http://emweb.eastmoney.com/PC_HSF10/NewFinanceAnalysis/zcfzbAjaxNew?companyType=4&reportDateType=0&" \
                  f"reportType=1&dates={year_list[0]}-{month_list[0]}-{month_day[str(month_list[0])]}%2C" \
                  f"{year_list[1]}-{month_list[1]}-{month_day[str(month_list[1])]}%2C" \
                  f"{year_list[2]}-{month_list[2]}-{month_day[str(month_list[2])]}%2C" \
                  f"{year_list[3]}-{month_list[3]}-{month_day[str(month_list[3])]}%2C" \
                  f"{year_list[4]}-{month_list[4]}-{month_day[str(month_list[4])]}" \
                  f"&code={stock_demo1}"







"""根据url,获得HTML文本"""
def getHtml(url):
    try:
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36"
        header = {
            "user-agent": user_agent
        }
        resp = requests.get(url, headers=header,timeout=40)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding
        return resp.text
    except:
        return ""


"""根据文件解析数据"""

# 提取股票名称
def getstockname(resp_text):
    """股票名称"""
    soup =BeautifulSoup(resp_text, "html.parser")
    title_list = soup.find("title").text.split("(")
    title = title_list[0]
    return title


# 资产负债表的数据信息



"""提取相关的数据"""

#




"""保存数据"""



"""运行主程序"""
def main():
    """用户输入"""
    # 股票代码
    stock_code = "002460"
    """编辑正确的url链接"""
    # 计算url

    url="http://emweb.eastmoney.com/PC_HSF10/NewFinanceAnalysis/Index?type=web&code=sz002460"
    resp_text = getHtml(url)
    """得到股票名字等信息"""
    stock_title = getstockname(resp_text)















































