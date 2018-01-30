# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import requests
from bs4 import BeautifulSoup
from config.Constant import *
import xlrd
import json

# init params
filename =  sys.path[0] + '\\' + 'config.properties'
p = Properties(filename)
properties = p.getProperties()
start_url = properties.get('start_url')
targer_url = properties.get('targer_url')
validate_code_url = properties.get('validate_code_url')
validate_code_file = properties.get('validate_code_file')
headers = {'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",'Referer':"https://sd.122.gov.cn/views/inquiry.html"}
cookies = {}
failed_carinfo = []

def get_validatecode():
    while True:
        # get captcha
        res = requests.get(validate_code_url,headers=headers,cookies=cookies)
        cookies.update(res.cookies)
        # print ('validate code_new_cookies:\t',res.cookies)
        # print ('validate code_current_public cookies:\t',cookies)
        f = open(validate_code_file+'temp.jpg', 'wb')  ##写入多媒体文件必须要 b 这个参数！！必须要！！
        f.write(res.content)  ##多媒体文件要是用conctent哦！
        f.close()
        # print ('please input validate_code from %s test.jpg , c=change\n' %(validate_code_file))
        message = str('请输入验证码（验证码位置——' + validate_code_file + '\\temp.jpg），输入\'c\'代表更换验证码，请输入：\n').decode('utf-8').encode('gbk')
        print ('%s' % (message))
        validate_code = raw_input()
        if validate_code != 'c':
            return validate_code

def query(c):

    validate_code = get_validatecode()
    message = str('请稍等...').decode('utf-8').encode('gbk')
    print ('%s'% (message))
    # send request
    res = requests.get(start_url,headers=headers,cookies=cookies)
    cookies.update(res.cookies)
    # print ('index_new_cookies:\t', res.cookies)
    # print ('index_current_public cookies:\t', cookies)
    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text,'lxml')
    # soup.original_encoding = 'utf-8'
    # get qm and page
    # print ----------------------------------------------------------------------------------------
    qm = soup.find_all('input',attrs={'name':'qm'})[0].attrs['value']
    page = soup.find_all('input',attrs={'name':'page'})[0].attrs['value']
    # get hpzl dict
    hpzl_dict = {}
    hpzl_list = soup.find('select',attrs={'name':'hpzl'}).find_all('option')
    for item in hpzl_list:
        # hpzl_dict[item.attrs['value']] = str(item.string).decode('unicode_escape')
        hpzl_dict[item.string] = item.attrs['value']
    # print hpzl_dict

    # name = c.name
    hphm1b = unicode(c.carnum[1:])
    hphm = unicode(c.carnum).encode('utf-8')
    enginnum = str(c.enginnum)
    if hpzl_dict.has_key(c.kind):
        hpzl = hpzl_dict.get(c.kind)
    else:
        hpzl = '01'
        # print 'default 01'
    payload = {
        'hpzl':hpzl,
        'hphm1b':hphm1b,
        'hphm':hphm.decode('utf-8'),
        'fdjh':enginnum[enginnum.__len__()-6:],
        'captcha':validate_code,
        'qm':qm,
        'page':page
    }
    # print payload

    res = requests.post(targer_url,payload,headers=headers,cookies=cookies)
    cookies.update(res.cookies)
    # print ('query URL_new_cookies:\t', res.cookies)
    # print ('query URL_current_public cookies:\t', cookies)
    # print res.content
    result_json = json.loads(res.content)

    if result_json['code'] == 200:
        result_str = '该机动车非现场未处理违法记录共计%d条。其中，牌证发放地违法记录%d条，跨省违法%d条。' % (result_json['data']['content']['zs'],result_json['data']['content']['bd'],result_json['data']['content']['ws'])
        return result_str
    elif result_json['code'] == 499:
        print result_json['message']
        return query(c)
    else:
        failed_carinfo.append(c)
        return result_json['message']

def prn_obj(obj):
    print '\n'.join(['%s:%s' % item for item in obj.__dict__.items()])

# excel
excelfilename = properties.get('excel')
# data = xlrd.open_workbook(r'E:\test.xlsx')
data = xlrd.open_workbook(excelfilename)
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols
carinfolist = []
for i in xrange(1,nrows):
    rowValues = table.row_values(i)
    carinfo = CarInfo(name=rowValues[0],carnum=rowValues[1],enginnum=rowValues[2],kind=rowValues[3])
    carinfolist.append(carinfo)
for item in carinfolist:
    result = query(item)
    print ('车辆<%s>查询结果：[%s]' %(item.name,result))
    print '********************************************************************************************************'
if failed_carinfo:
    message = str('以下是查询失败的车辆信息，请手动查询：').decode('utf-8').encode('gbk')
    print ('%s' % (message))
    for item in failed_carinfo:
        prn_obj(item)
message = str('按Enter结束...').decode('utf-8').encode('gbk')
print ('%s' % (message))
raw_input('')



