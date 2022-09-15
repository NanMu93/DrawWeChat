import re
import time
from appium import webdriver
import pandas as pd
import numpy as np
import openpyxl
from appium.webdriver.common.touch_action import TouchAction
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from config import *



class Moments():
    def __init__(self):
        """
        初始化
        """
        # 驱动配置
        self.desired_caps = {
            'platformName': PLATFORM,
            'deviceName': DEVICE_NAME,
            'appPackage': APP_PACKAGE,
            'appActivity': APP_ACTIVITY,
            'noReset': NORESET
        }
        self.driver = webdriver.Remote(DRIVER_SERVER, self.desired_caps)
        self.wait = WebDriverWait(self.driver, TIMEOUT)

    def fileTrans(self):
        """
        查找文件传输助手
        :return:
        """
        filetrans = self.wait.until(EC.presence_of_element_located((By.XPATH, '//android.view.View[@text="文件传输助手"]')))
        filetrans.click()
        sleep(SCROLL_SLEEP_TIME)

    def contact(self):
        """
        定位通讯录
        :return:
        """
#        tab = self.wait.until(EC.presence_of_element_located((By.XPATH, '//android.widget.ImageView[@resource-id="com.tencent.mm:id/f2s"][3]')))
        tab = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@resource-id="com.tencent.mm:id/f2a"]')))[1]
        tab.click()
        # 找到目标
#        sleep(SCROLL_SLEEP_TIME)
        while True:
            try:
                selectf = self.driver.find_element(By.XPATH, '//*[@text="' + FRIEND + '"]')
                selectf.click()
                break
            except NoSuchElementException:
                self.driver.swipe(FLICK_START_X, FLICK_START_Y + 1000, FLICK_START_X, FLICK_START_Y, 1000)
                pass
#        selectf = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@text="' + FRIEND + '"]')))
#        selectf.click()
        # 进入朋友圈
        f = self.wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@resource-id="com.tencent.mm:id/iwg"][3]')))
        f.click()

    def craw(self):
        """
        抓取
        :return:
        """

        #        df = pd.read_excel("venv/data/wechat.xlsx")

        # 滑动到结尾
        end_point1 = "com.tencent.mm:id/g39"
        end_point2 = "com.tencent.mm:id/ifi"
        end_point3 = False
        # 年份元素
        years = "com.tencent.mm:id/jxl"
        # 时间元素
        wdate = "com.tencent.mm:id/ju9"
        # 月份元素
        mdate = "com.tencent.mm:id/juc"
        # 日期元素
        ddate = "com.tencent.mm:id/jsu"
        # 含图文字内容
        psms = "com.tencent.mm:id/c22"
        # 含链接文字内容和纯文字内容
        wsms = "com.tencent.mm:id/c2h"
        # 链接标题
        llink = "com.tencent.mm:id/kpq"
        ryear_list = []
        rdate_list = []
        rtext_list = []
        rlink_list = []
        dict1 = {}
        tyear = str(time.localtime().tm_year) + "年"
        rdate = ' '
        if len(TDATETIME) == 0:
            ttime = time.mktime(time.strptime("1900年1月1日", "%Y年%m月%d日"))
        else:
            ttime = time.mktime(time.strptime(TDATETIME, "%Y年%m月%d日"))
        print(ttime)
        sleep(SCROLL_SLEEP_TIME)
        while True:
            # 当前页面显示的所有目标内容集合
            #            items = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@resource-id="com.tencent.mm:id/br8"]')))
            items = self.wait.until(EC.presence_of_all_elements_located((By.ID, "com.tencent.mm:id/br8")))
            # 遍历每条内容
            for item in items:
                rtext = ' '
                ryear = ' '
                rlink = ' '
                try:
                    ryear = item.find_element(By.ID, years).get_attribute('text')
                    dict2 = {}
                    dict2['年份'] = ryear_list
                    dict2['日期'] = rdate_list
                    dict2['内容'] = rtext_list
                    dict2['链接'] = rlink_list
                    df2 = pd.DataFrame(dict2)
                    with pd.ExcelWriter('F:/test.xlsx', mode='a', engine='openpyxl') as writer:
                        writer.if_sheet_exists = "replace"
                        df1 = pd.read_excel('F:/test.xlsx', sheet_name='Sheet1', index_col=0)
                        f = [df1, df2]
                        result = pd.concat(f, axis=0)
                        result.to_excel(writer, sheet_name="Sheet1", index_label=0)
                    ryear_list = []
                    rdate_list = []
                    rtext_list = []
                    rlink_list = []
                    tyear = ryear
                #    print(ryear + "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                except NoSuchElementException:
                    pass
                try:
                    rtext = item.find_element(By.ID, psms).get_attribute('text')
                    try:
                        a = rtext_list.index(rtext)
                        continue
                    except ValueError:
                        pass
                except NoSuchElementException:
                    pass
                try:
                    rtext = item.find_element(By.ID, wsms).get_attribute('text')
                    try:
                        a = rtext_list.index(rtext)
                        continue
                    except ValueError:
                        pass
                except NoSuchElementException:
                    pass
                try:
                    rlink = item.find_element(By.ID, llink).get_attribute('text')
                    try:
                        a = rtext_list.index(rlink)
                        continue
                    except ValueError:
                        pass
                except NoSuchElementException:
                    pass
                try:
                    rdate = item.find_element(By.ID, ddate).get_attribute('text')
                    try:
                        rdate1 = item.find_element(By.ID, mdate).get_attribute('text')
                        rdate = rdate1 + rdate + "日"
                        ntime = time.mktime(time.strptime(tyear + rdate, "%Y年%m月%d日"))
                        print(ntime)
                        if ntime < ttime:
                            end_point3 = True
                            break
                    except NoSuchElementException:
                        print("没取到月份")
                        pass
                except NoSuchElementException:
                    pass
                rdate_list.append(rdate)
                rlink_list.append(rlink)
                rtext_list.append(rtext)
                ryear_list.append(ryear)
            bounds = items[len(items) - 1].get_attribute('bounds')
            m = re.findall(r'\d+', bounds)
            print(m)
            print(type(m))
            half_m = int(m[len(m) - 1]) / 2 - 100
            # 判断是否已经到结尾
            page = self.driver.page_source
            if end_point1 in page or end_point2 in page or end_point3:
                break
            # 上滑
            self.driver.swipe(FLICK_START_X, FLICK_START_Y + half_m, FLICK_START_X, FLICK_START_Y, 2000)
            self.driver.swipe(FLICK_START_X, FLICK_START_Y + half_m, FLICK_START_X, FLICK_START_Y, 2000)
        dict1['年份'] = ryear_list
        dict1['日期'] = rdate_list
        dict1['内容'] = rtext_list
        dict1['链接'] = rlink_list
        print(dict1)
        df = pd.DataFrame(dict1)
        #    df.to_excel('venv/data/wechat.xlsx', sheet_name='sheet1')
        with pd.ExcelWriter('F:/test.xlsx', mode='a', engine='openpyxl') as writer:
            writer.if_sheet_exists = "replace"
            df1 = pd.read_excel('F:/test.xlsx', sheet_name='Sheet1', index_col=0)
            f = [df1, df]
            result = pd.concat(f, axis=0)
            result.to_excel(writer, sheet_name="Sheet1", index_label=0)

    def main(self):
        """
        入口
        :return:
        """
        # 通讯录
        self.contact()
        # 爬取
        self.craw()


if __name__ == '__main__':
    moments = Moments()
    moments.main()
