#- * - coding: utf - 8 -*-

'''
    author:Kevin
    date:2019/1/24
'''

from ctypes import *
from PIL import Image,ImageEnhance
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from ctypes import *
import pytesseract
import os
import sys
import time
import re


if __name__ == '__main__':


    # 高级检索
    # 注意：高级检索只能检索 年份*一个检索方式，若需要多项筛选请移步下面的专业检索
    # 从上到下依次为：检索方式、起止日期、截止日期、作者单位和页码
    add_condition = "#add_condition > i"
    search_method = "作者单位"
    publishdate_from = "2018年"#“年”是必须写的，与选项中保持一致
    publishdate_to = "2019年" #比如选项中有“至今”这个选项，直接把参数改为至今就可以
    author_school = "同济大学"
    page_number = "#laypage_8 > a:nth-child(2)"  #如果需要从相应页码开始爬取，则修改括号中的内容

    #专业检索
    advanced_search = "作者单位:(同济大学)*期刊名称/刊名:(SCI)"  #具体使用方法请参考网站

    driver = webdriver.Chrome('D:\\chromeDirver\\chromedriver.exe')
    driver.get('http://www.wanfangdata.com.cn/searchResult/getAdvancedSearch.do?searchType=all')

    #最大化窗口
    #driver.maximize_window()

    #选择专业搜索
    #driver.find_element_by_css_selector("#expert_search_a").click()

    #assert "万方" in driver.title

    time.sleep(1)
    def selc(id,text):
        sele = driver.find_element_by_id(id)
        selc_1 = Select(sele)
        selc_1.select_by_visible_text(text)

    def selc_1(selector,text):
        sele = driver.find_element_by_css_selector(selector)
        selc_2 = Select(sele)
        selc_2.select_by_visible_text(text)

    def check_all():
        chk_all = (By.NAME,"checkAll")
        if WebDriverWait(driver, 20, 0.5).until(EC.presence_of_element_located(chk_all)):
            driver.find_element_by_name("checkAll").click()
        else:
            print("请重新启动脚本")

    def next_page():
        nex_page = (By.CLASS_NAME,"laypage_next")
        if WebDriverWait(driver, 20, 0.5).until(EC.presence_of_element_located(nex_page)):
            driver.find_element_by_class_name("laypage_next").click()
        else:
            print("请重新启动脚本")

    def is_element_exist(className):
        try:
            driver.find_element_by_class_name(className)
            return True
        except:
            return False

    def is_element_exist_id(id):
        try:
            driver.find_element_by_class_name(id)
            return True
        except:
            return False


    def export():
        driver.find_element_by_css_selector(
            "#search_result_2 > div.BatchOper > a:nth-child(3)").click()
        for handle in driver.window_handles:
            driver.switch_to_window(handle)
        time.sleep(5)
        driver.find_element_by_xpath("//*[@id='tags']/li[5]/a").click()
        time.sleep(5)
        driver.find_element_by_name("exportBtn").click()
        time.sleep(3)
        # 跳转主页面
        windows = driver.window_handles
        driver.switch_to.window(windows[0])
        driver.find_element_by_xpath("//*[@id='search_result_2']/div[2]/a[2]").click()

    time.sleep(3)


    if is_element_exist_id("expert_search_textarea"):
        print("已选择专业搜索")
        driver.find_element_by_css_selector("#expert_search_textarea").send_keys(advanced_search)
        elem_search_a = driver.find_element_by_css_selector("#ch_button")
        elem_search_a.click()
        elem_search_a.send_keys(Keys.RETURN)
        time.sleep(10)
    else:
        selc("gaoji",search_method)
        driver.find_element_by_css_selector("#ddd").send_keys(author_school)
        selc("advanced_search_publshdate_start",publishdate_from)
        selc("advanced_search_publshdate_end",publishdate_to)
        elem_search = driver.find_element_by_css_selector("#set_advanced_search_btn")
        elem_search.click()
        elem_search.send_keys(Keys.RETURN)
        time.sleep(10)


    selc("select4", "每页显示50条")
    #若只显示文章名字，则取消注释下面这条语句
    # driver.find_element_by_css_selector("#icon_icon_menu2_2").click()

    count = 0
    #!!!尝试每选取一页就导出文件，避免错误
    # while is_element_exist("laypage_next"):
    #     time.sleep(30)
    #     check_all()
    #     message = driver.find_element_by_css_selector("#search_result_2 > div.BatchOper > a:nth-child(1)").text
    #     print(message)
    #     x = int(str(message).strip("全选()")) #记录页数
    #
    #     if  x%50 == 0:
    #         print("正在爬取第%s页。" % ((x/50)+(count*10)))
    #     else:
    #         print("爬取发生错误第%s页，有漏选论文，请重新人工检查,或者重新启动。" % ((x/50)+(count*10)))
    #         print("由于错误将提前导出文件")
    #         time.sleep(1)
    #         export()
    #
    #     while message == "全选(500)":
    #         count = count + 1
    #         time.sleep(1)
    #         export()
    #         time.sleep(1)
    #         print("第%s次下载完成（每次下载10页，共500条）" % count)
    #
    #     time.sleep(30)
    #     next_page()
    ex = 1
    while is_element_exist("laypage_next"):
            time.sleep(15)
            check_all()
            print("-------------------------------------")
            print("正在爬取第%s" % ex)
            message = driver.find_element_by_css_selector("#search_result_2 > div.BatchOper > a:nth-child(1)").text
            x = int(str(message).strip("全选()"))  # 记录页数
            if x%50 == 0:
                print("爬取正常！")
            else:
                print("当前页码爬取出错，请检查第%s页" % count)
            export()
            time.sleep(1)
            print("第%s次下载完成" % ex)
            print("-------------------------------------")
            time.sleep(2)
            next_page()
            ex = ex + 1

    print("爬取结束，请工作人仔细核对文件正确与否")
    print("最后一部分需要人工导出文件")