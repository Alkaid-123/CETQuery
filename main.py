from time import sleep
import requests
import pandas as pd

data = pd.read_excel('data.xls')

data.insert(2, '准考证号', '')
data.insert(3, '学校', '')
data.insert(4, '四级成绩', '')
data.insert(5, '听力部分4', '')
data.insert(6, '阅读部分4', '')
data.insert(7, '写作部分4', '')
data.insert(8, '六级成绩', '')
data.insert(9, '听力部分6', '')
data.insert(10, '阅读部分6', '')
data.insert(11, '写作部分6', '')

headers = {
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive',
    'Host': 'cachecloud.neea.cn',
    'Origin': 'http://cet.neea.cn',
    'Referer': 'http://cet.neea.cn/',
    'Cookie': 'Hm_lvt_dc1d69ab90346d48ee02f18510292577=1629958972,1629961007,1629989056,1629992054; community=Home; '
              'language=1; http_waf_cookie=f27d75db-23f7-4aa938445a95fb8092bcad0ba73630744bd3',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/98.0.4758.102 Safari/537.36 '
}

params_list = []

for i in range(len(data)):
    # 四级
    params_list.append({
        'km': 1,
        'xm': data.iloc[i, 0],
        'no': data.iloc[i, 1],
        'source': 'pc'
    })
    # 六级
    params_list.append({
        'km': 2,
        'xm': data.iloc[i, 0],
        'no': data.iloc[i, 1],
        'source': 'pc'
    })

url = "http://cachecloud.neea.cn/api/latest/results/cet"
print(params_list)

for i in range(len(params_list)):
    response = requests.get(url, params_list[i], headers=headers)
    if i % 2 == 0:
        print("开始爬取%s的四级成绩......" % params_list[i]['xm'])
    else:
        print("开始爬取%s的六级成绩......" % params_list[i]['xm'])
    plus = (i % 2) * 4
    try:
        json_data = response.json()
        idx = i // 2
        data.iloc[idx, 2] = json_data['zkzh']
        data.iloc[idx, 3] = json_data['xx']
        data.iloc[idx, 4 + plus] = json_data['score']
        data.iloc[idx, 5 + plus] = json_data['sco_lc']
        data.iloc[idx, 6 + plus] = json_data['sco_rd']
        data.iloc[idx, 7 + plus] = json_data['sco_wt']
        info = '%d：%s查询成功，成绩为：%s' % ((i + 1), params_list[i]['xm'], json_data['score'])
        print(info)
    except Exception as e:
        print('%d：%s未参加考试...' % ((i + 1), params_list[i]['xm']))
    print("*******************************************************************")

    sleep(0.3)

print("已全部查询完毕，正在导出到csv......")
data.to_csv('grade.csv', encoding='gbk')
print("导出完毕.")