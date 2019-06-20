# -*- coding: utf-8 -*-
'''
 @Time    : 2019/6/11 9:37
 @Author  : YangHao
 @File    : YaMaXun.py
 @Software: PyCharm
 '''


import win32com
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
from selenium import webdriver
from lxml import etree
import time
import csv
import string #引入字符模块
import random
import selenium.webdriver.support.expected_conditions as EC


# 8位用户名及6位密码
def get_userNameAndPassword():
    global userName, password,dian
    dian=int(random.randint(2, 11))
    L1 = string.ascii_lowercase  # 所有的小写字母
    L2 = string.ascii_uppercase  # 所有的大写字母
    L3 = string.digits  # 所有的数字
    # L4 = string.punctuation #所有的特殊字符

    password = random.sample(L1, 2) + random.sample(L2, 1) + random.sample(L3, 5)
    userName = random.sample(L1, 2) + random.sample(L2, 2) + random.sample(L3, 4)
    password = ''.join(password)
    userName = ''.join(userName)
#写入文件
def csv_login():
    with open('login_620.csv', 'a+', encoding='utf-8-sig', newline='') as csvfile:  # 文件后缀.csv
        csvObj = csv.writer(csvfile)  # 声明 csv 对象
        a = [email, password]
        csvObj.writerow(a)  # 写入数据
        pass
#爬取地址信息
def get_add(bro):
    # 需要获取 街道地址 公寓号 市 州 邮编 电话号码
    # url_add = 'http://www.haoweichi.com/'
    # bro = webdriver.Chrome()
    # bro.maximize_window()
    # bro.set_page_load_timeout(20)
    # bro.get(url=url_add)
    global jiedao,shi,zhou,youbian,phone,idName,idNumber,period
    bro.execute_script(
        'window.open("http://www.haoweichi.com/")')
    handles = bro.window_handles
    bro.implicitly_wait(10)  # 隐示等待
    bro.switch_to.window(bro.window_handles[2])
    time.sleep(5)
    print('获取地址中........')
    jiedao = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[6]/div[2]/input').get_attribute(
        'value')  # 获取到街道地址
    shi = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[7]/div[2]/input').get_attribute(
        'value')  # 获取到市
    zhou = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[5]/div[4]/input').get_attribute(
        'value')  # 获取到州
    youbian = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[8]/div[2]/input').get_attribute(
        'value')  # 获取到邮编
    phone = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[7]/div[4]/input').get_attribute(
        'value')  # 获取到电话号
    #/html/body/div[1]/div[3]/div[2]/div/div[2]/div[2]/div[2]/input
    idName = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[2]/div[2]/input').get_attribute(
        'value')  # 获取到卡
    idNumber = bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[12]/div[4]/input').get_attribute(
        'value')  # 获取到电话号
    period=bro.find_element_by_xpath('/html/body/div[1]/div[3]/div[2]/div/div[2]/div[13]/div[4]/input').get_attribute('value')
    time.sleep(5)

    print('jiedao：',jiedao)
    print('shi：',shi)
    print('zhou：',zhou)
    print('zip：',youbian)
    print('phone：',phone)
    print('idName：', idName)
    print('idNumber：', idNumber)
    print('period：', period)
    pass
#地址填写
def write_add(bro,addUrl,shopurl):
    #https://www.amazon.com/Value-CCELL-Model-Cartridges-Ceramic/dp/B07Q9XXLQF/ref=mp_s_a_1_14?keywords=ccell+cartridge&qid=1560405846&s=gateway&sr=8-14
    bro.execute_script(
        'window.open("%s")'%addUrl)
    handles = bro.window_handles
    time.sleep(5)
    bro.switch_to.window(bro.window_handles[3])
    time.sleep(5)
    bro.find_element_by_xpath('//*[@id="contextualIngressPtLabel"]').click()
    time.sleep(10)
    #//*[@id="olpOfferList"]/div/div/div[3]/div[4]/h3/span/a
    bro.find_element_by_xpath('//*[@id="GLUXZipUpdateInput"]').clear()
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="GLUXZipUpdateInput"]').send_keys(youbian)
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="GLUXZipUpdate"]/span/input').click()
    time.sleep(5)
    #..........................//*[@id="a-popover-6"]/div/div[2]/span
    bro.find_element_by_xpath('//*[@class="a-popover-wrapper"]/div[2]/span').click()
    time.sleep(5)
    bro.find_element_by_xpath('//*[@id="olp-upd-new-freeshipping"]/span/a/b').click()
    time.sleep(5)
    shoppAll = BeautifulSoup(bro.page_source)
    div=shoppAll.find_all('div',class_='a-row a-spacing-mini olpOffer')
    itemget = []
    for item in div:
        item=item.find('span',class_='a-size-medium a-text-bold').a.string
        itemget.append(item)
    i=0
    for n in  itemget:
        if n==shopurl:#//*[@id="a-autoid-0"]/span/input
            bro.find_element_by_xpath('//*[@id="a-autoid-%d"]/span/input' % i).click()
        i = i + 1
    ####添加购物车
    bro.implicitly_wait(5)  # 隐示等待
    bro.find_element_by_xpath('//*[@id="hlb-ptc-btn-native"]').click()
    bro.implicitly_wait(5)  # 隐示等待
    #=====================================================================
    print('address..........')
    bro.find_element_by_xpath('//*[@id="enterAddressAddressLine1"]').clear()
    time.sleep(1)
    # 光标移动到街道地址框
    bro.find_element_by_xpath('//*[@id="enterAddressAddressLine1"]').send_keys(jiedao+str(' c/o'))
    # 输入街道地址
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="enterAddressAddressLine2"]').clear()
    time.sleep(1)
    # 光标移动到街道地址框
    bro.find_element_by_xpath('//*[@id="enterAddressAddressLine2"]').send_keys(jiedao + str(' c/o'))
    # 输入街道地址
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="enterAddressCity"]').clear()
    time.sleep(1)
    # 光标移动到市
    bro.find_element_by_xpath('//*[@id="enterAddressCity"]').send_keys(shi)
    # 输入市
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="enterAddressStateOrRegion"]').clear()
    time.sleep(1)
    # 光标移动到州
    bro.find_element_by_xpath('//*[@id="enterAddressStateOrRegion"]').send_keys(zhou)
    # 输入州
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="enterAddressPostalCode"]').clear()
    time.sleep(1)
    # 光标移动到邮编
    bro.find_element_by_xpath('//*[@id="enterAddressPostalCode"]').send_keys(youbian)
    # 输入邮编
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="enterAddressPhoneNumber"]').clear()
    time.sleep(1)
    # 光标移动到电话号码
    bro.find_element_by_xpath('//*[@id="enterAddressPhoneNumber"]').send_keys(phone)
    # 输入电话号码
    time.sleep(1)
    # 添加地址
    bro.find_element_by_xpath('//*[@id="newShippingAddressFormFromIdentity"]/div[1]/div/form/div[6]/label[2]/input').click()
    time.sleep(1)

    bro.find_element_by_xpath('//*[@id="newShippingAddressFormFromIdentity"]/div[1]/div/form/div[7]/span').click()
    time.sleep(10)  # 隐示等待
    bro.find_element_by_xpath('//*[@id="newShippingAddressFormFromIdentity"]/div[1]/div/form/div[7]/span').click()
    bro.implicitly_wait(5)  # 隐示等待
    # =====================================================================
    #确认是否继续
    bro.find_element_by_xpath('//*[@id="shippingOptionFormId"]/div[1]/div[2]/div/span[1]').click()
    time.sleep(10)
    #添加信用卡]div[7]/div/div[1]/div[2]/div[3]/div[1]/div[1]/div[0]/div[]

    while True:

        try:
            bro.find_element_by_xpath('//form/div[3]/div[1]/input').clear()
            time.sleep(3)
            bro.find_element_by_xpath('//form/div[3]/div[1]/input').send_keys(idName)
            time.sleep(3)   #/form/div[3]/div[2]/div/
            idA=bro.find_element_by_xpath('//*[@name="addCreditCardNumber"]')
            time.sleep(1)
            idA.clear()
            time.sleep(1)
            idA.send_keys(idNumber)
            break
        except Exception as e:
            print('No xpath')
            time.sleep(10)
            bro.refresh()
            rev=0
            if rev==3:
                break
            rev+=1

    month, year =period.split('/')
    month = month.split('0')
    if len(month) == 2:
        month = month[1]
    else:
        month=month[0]  #//*[@id="form-add-credit-card"]/div[3]/div[4]/span[1]/span  //*[@id="pp-38-48"]/div[3]/div[4]/div/span[1]
    time.sleep(5)
    T=bro.find_element_by_xpath('//form/div[3]/div[4]/div[1]/span[1]')
    if T:
        T.click()
        time.sleep(1)
    else:
        bro.find_element_by_xpath('//form/div[3]/div[4]/span[1]').click()
        time.sleep(1)
    time.sleep(2)
    bro.find_element_by_xpath('//*[@id="1_dropdown_combobox"]/li[%s]/a' % month).click()
    time.sleep(2)
    Y=bro.find_element_by_xpath('//form/div[3]/div[4]/div[1]/span[3]')
    if Y:
        Y.click()
        time.sleep(1)
    else:
        bro.find_element_by_xpath('//form/div[3]/div[4]/span[3]').click()
        time.sleep(1)
    if year=='2019':
        bro.find_element_by_xpath('///*[@id="2_dropdown_combobox"]/li[1]/a').click()
        time.sleep(1)
    elif year =='2020':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[2]/a').click()
        time.sleep(1)
    elif year == '2021':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[3]/a').click()
        time.sleep(1)
    elif year == '2022':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[4]/a').click()
        time.sleep(1)
    elif year == '2023':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[5]/a').click()
        time.sleep(1)
    elif year == '2024':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[6]/a').click()
        time.sleep(1)
    elif year == '2025':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[7]/a').click()
        time.sleep(1)
    elif year == '2026':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[8]/a').click()
        time.sleep(1)
    elif year == '2027':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[9]/a').click()
        time.sleep(1)
    elif year == '2028':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[10]/a').click()
        time.sleep(1)
    elif year == '2029':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[11]/a').click()
        time.sleep(1)
    elif year == '2030':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[12]/a').click()
        time.sleep(1)
    elif year == '2031':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[13]/a').click()
        time.sleep(1)
    elif year == '2032':
        bro.find_element_by_xpath('//*[@id="2_dropdown_combobox"]/li[14]/a').click()
        time.sleep(1)
    time.sleep(2)
        #//*[@id="pp-Hf-58"]/span
    bro.find_element_by_xpath('//form/div[3]/div[5]/span/span').click()
    #//*[@id="pp-YC-51"]/span/input

    ka=bro.find_element_by_xpath('//form/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]'
                              '/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/label/input')
    if ka:
        ka.click()
    bro.find_element_by_xpath('//form/div[2]/div[1]/div[1]/div[1]/span[1]/span[1]').click()
    time.sleep(5)
    bro.find_element_by_xpath('//*[@id="address-book-entry-0"]/div[2]/span/a').click()
    time.sleep(5)

    pass
# 注册账户
def ymx_namepwd(addUrl,shopurl):


    try:
        get_userNameAndPassword()
        print("userName:", userName)
        print("Password:", password)
    except Exception as e:
        print(e.args)
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"')  # 设置请求头的User-Agent
    chrome_options.add_argument('--disable-infobars')  # 禁用浏览器正在被自动化程序控制的提示
    # chrome_options.add_argument('--incognito')  # 隐身模式（无痕模式）
    # chrome_options.add_argument('--headless')#浏览器不提供可视化页面
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    # prefs = {"profile.managed_default_content_settings.images": 2}
    # chrome_options.add_experimental_option("prefs", prefs)#禁止浏览器加载图片  提升爬虫效率
    url1 = 'http://www.linshiyouxiang.net/'
    # chrome_options = chrome_options
    bro = webdriver.Chrome(chrome_options = chrome_options)
    bro.maximize_window()
    bro.set_page_load_timeout(20)
    bro.get(url=url1)
    bro.implicitly_wait(5)  # 隐示等待
    # // *[ @ id = "top"] / div / div / div[2] / div / div[2] / button
    bro.find_element_by_xpath('// *[ @ id = "top"] / div / div / div[2] / div / div[2] / button').click()
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="top"]/div/div/div[2]/div/div[2]/ul/li[%s]/a'%dian).click()
    time.sleep(1)
    global email
    email = bro.find_element_by_xpath('//input[@id="active-mail"]')
    email = email.get_attribute('data-clipboard-text')  # 获取到email地址
    print('email:', email)
    # =======================================================
    bro.execute_script(
        'window.open("https://www.amazon.com/ap/register?_encoding=UTF8&openid.assoc_handle=usflex&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.mode=checkid_setup&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&openid.ns.pape=http%3A%2F%2Fspecs.openid.net%2Fextensions%2Fpape%2F1.0&openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com%2Fgp%2Fyourstore%2Fhome%3Fie%3DUTF8%26ref_%3Dnav_newcust")')
    handles = bro.window_handles
    bro.implicitly_wait(5)  # 隐示等待
    bro.switch_to.window(bro.window_handles[1])
    time.sleep(2)
    bro.find_element_by_xpath('//*[@id="ap_customer_name"]').clear()
    time.sleep(1)
    # 光标移动到输入框
    bro.find_element_by_xpath('//*[@id="ap_customer_name"]').send_keys(userName)
    # 输入名字
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="ap_email"]').clear()
    # 光标移动到输入框
    bro.find_element_by_xpath('//*[@id="ap_email"]').send_keys(email)
    # 输入邮箱
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="ap_password"]').clear()
    time.sleep(1)
    # 光标移动到输入框
    bro.find_element_by_xpath('//*[@id="ap_password"]').send_keys(password)
    # 输入密码
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="ap_password_check"]').clear()
    time.sleep(1)
    # 光标移动到输入框
    bro.find_element_by_xpath('//*[@id="ap_password_check"]').send_keys(password)
    # 再次输入密码
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="continue"]').click()
    time.sleep(1)
    bro.switch_to.window(bro.window_handles[0])
    time.sleep(3)
    w = 0
    while True:
        try:
            a = bro.find_element_by_xpath('//*[@id="message-list"]/tr[1]/td[3]/a')  # 获取a标签

            if a:  # 如果收到邮件  a标签则有
                a.click()  # 点击a标签
                bro.implicitly_wait(5)  # 隐示等待
                html = bro.page_source  # 获取点击之后的源码
                html = etree.HTML(html)  # 解析
                yanzhengma = html.xpath('//*[@id="verificationMsg"]/p[2]/text()')  # 解析寻找验证码，找到xpath路径解析一下就好了
                print('yanzhengma:', yanzhengma)
                break
        except Exception as e:
            print('No email,please wait')
            time.sleep(10)
    time.sleep(3)
    bro.switch_to.window(bro.window_handles[1])
    time.sleep(3)
    bro.find_element_by_xpath('//*[@id="cvf-page-content"]/div/div/div[1]/form/div[2]/input').clear()
    time.sleep(1)
    # 光标移动到输入框
    bro.find_element_by_xpath('//*[@id="cvf-page-content"]/div/div/div[1]/form/div[2]/input').send_keys(yanzhengma)
    # 输入验证码
    time.sleep(1)
    bro.find_element_by_xpath('//*[@id="a-autoid-0"]/span/input').click()
    time.sleep(5)
    # 文件写入
    csv_login()
    #获取收货地址
    get_add(bro)
    # 收货地址填写
    write_add(bro,addUrl,shopurl)
    #关闭窗口
    bro.close()
    bro.quit()
    pass

if __name__ == '__main__':

    # add_list = ['https://www.amazon.com/Value-CCELL-Model-Cartridges-Ceramic/dp/B07Q9XXLQF/ref=mp_s_a_1_14?keywords=ccell+cartridge&qid=1560405846&s=gateway&sr=8-14']
    # list2 = ['MJD Sales']
    # for addUrl in add_list:
    #     for shopurl in list2:

    while True:
        y = input('input(y or n)：')
        if y == 'y':
            shopName=input('Shop Name:')
            addUrl=input('Commodity Links:\n')
            try:
                ymx_namepwd(addUrl,shopName)
            except Exception as e:
                print('procedure runs error', e.args)
                time.sleep(1)
        else:
            print('End of Program Running ')
            break






