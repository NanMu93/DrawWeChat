import time
from datetime import datetime

print(time.localtime(time.time()).tm_year)
print(type(time.localtime().tm_year))
print(datetime.strptime("2022年" + "9月10日", "%Y年%m月%d日"))
print(time.mktime(time.localtime()))
print(time.mktime(time.strptime("2022年" + "9月10日", "%Y年%m月%d日")))


    def craw(self):
        """
        抓取
        :return:
        """

#        df = pd.read_excel("venv/data/wechat.xlsx")

        # 滑动到结尾
        end_point1 = "com.tencent.mm:id/g39"
        end_point2 = "com.tencent.mm:id/ifi"
        # 年份元素
        years = "com.tencent.mm:id/jxl"
        # 时间元素
        wdate = "com.tencent.mm:id/ju9"
        # 月
        mdate = "com.tencent.mm:id/juc"
        # 日
        ddate = "com.tencent.mm:id/jsu"
        # 图片稿
        psms = "com.tencent.mm:id/c22"
        # 链接和纯文字
        wsms = "com.tencent.mm:id/c2h"
        # 链接文字
        llink = "com.tencent.mm:id/kpq"
        ryear_list = []
        rdate_list = []
        rtext_list = []
        rlink_list = []
        dict1 = {}
        sleep(SCROLL_SLEEP_TIME)
        while True:
            # 当前页面显示的所有状态
#            items = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@resource-id="com.tencent.mm:id/br8"]')))
            items = self.wait.until(EC.presence_of_all_elements_located((By.ID, "com.tencent.mm:id/br8")))
            # 遍历每条状态
            for item in items:
                rtext = ' '
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
                rtext_list.append(rtext)
                try:
                    rlink = item.find_element(By.ID, llink).get_attribute('text')
                    try:
                        a = rtext_list.index(rlink)
                        continue
                    except ValueError:
                        pass
                    rlink_list.append(rlink)
                except NoSuchElementException:
                    rlink_list.append(' ')
                    pass
                try:
                    ryear = item.find_element(By.ID, years).get_attribute('text')
                    ryear_list.append(ryear)
                except NoSuchElementException:
                    ryear_list.append(' ')
                    pass
                try:
                    rdate = item.find_element(By.ID, ddate).get_attribute('text')
                    try:
                        rdate1 = item.find_element(By.ID, mdate).get_attribute('text')
                        rdate = rdate1 + rdate
                        rdate_list.append(rdate)
                    except NoSuchElementException:
                        rdate_list.append(rdate)
                        pass
                except NoSuchElementException:
                    rdate_list.append(rdate)
                    pass
            bounds = items[len(items) - 1].get_attribute('bounds')
            print(bounds)
            m = re.findall(r'\d+', bounds)
            print(m)
            print(type(m))
            half_m = int(m[len(m) - 1]) / 2
            # 判断是否已经到结尾
            page = self.driver.page_source
            if end_point1 in page or end_point2 in page:
                break
            # 上滑
            self.driver.swipe(FLICK_START_X, FLICK_START_Y + half_m - 100, FLICK_START_X, FLICK_START_Y, 2000)
            self.driver.swipe(FLICK_START_X, FLICK_START_Y + half_m - 100, FLICK_START_X, FLICK_START_Y, 2000)
        dict1['年份'] = ryear_list
        dict1['日期'] = rdate_list
        dict1['内容'] = rtext_list
        dict1['链接'] = rlink_list
        print(dict1)
        df = pd.DataFrame(dict1)
        #    df.to_excel('venv/data/wechat.xlsx', sheet_name='sheet1')
        with pd.ExcelWriter('venv/data/wechat1.xlsx', mode='a', engine='openpyxl') as writer:
            writer.if_sheet_exists = "replace"
            df1 = pd.read_excel('venv/data/wechat1.xlsx', sheet_name='Sheet1', index_col=0)
            f = [df1, df]
            result = pd.concat(f, axis=0)
            result.to_excel(writer, sheet_name="Sheet1", index_label=0)
