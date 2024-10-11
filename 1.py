import requests
from bs4 import BeautifulSoup
import openpyxl

def get_weibo_cookie():
    url = 'https://passport.weibo.com/visitor/genvisitor2'
    payload = 'cb=visitor_gray_callback&tid=&from=weibo'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.post(url, data=payload, headers=headers, verify=False)
    return response.cookies.get_dict()['SUB']

def get_weibo_response():
    url = 'https://s.weibo.com/top/summary'
    cookie_str = get_weibo_cookie()
    cookie_dict = {'SUB': '='.join(cookie_str)}
    response = requests.get(url, cookies=cookie_dict)
    return response

def get_weibo_hot_search():
    response = get_weibo_response()
    soup = BeautifulSoup(response.text, 'html.parser')
    hot_trs = soup.find_all('tr')
    data = []
    for tr in hot_trs:
        a_tag = tr.find('a', href=True)
        if a_tag:
            title = a_tag.get_text(strip=True)
            href = a_tag['href']
            hotness_td = tr.find('td', class_='td-03')
            if hotness_td.get_text(strip=True) == '':
                hotness = '未知热度'
            else:
                hotness = hotness_td.get_text(strip=True)
            data.append({'title': title, 'href': href, 'hotness': hotness})
    return data

def print_weibo_data(data):
    for item in data:
        print(f'标题：{item["title"]}\n链接:{item["href"]}\n热度:{item["hotness"]}\n')

def write_to_excel(data, filename='weibo_hot_search.xlsx'):
    try:
        # 尝试加载现有的Excel文件
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
    except FileNotFoundError:
        # 如果文件不存在
        wb = openpyxl.Workbook()
        ws = wb.active
        # 添加表头
        ws.append(['标题', '链接', '热度'])

    for item in data:
        ws.append([item['title'], item['href'], item['hotness']])
    wb.save(filename)

def main():
    data = get_weibo_hot_search()
    # print_weibo_data(data)
    write_to_excel(data)

if __name__ == '__main__':
    main()