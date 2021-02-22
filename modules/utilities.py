from modules.libraries import *
from modules.globals import *

def getProdList():
    try:  # Mac
        wb = load_workbook(os.path.split(os.path.realpath(__file__))[0] + '/自動查最低價.xlsm')
    except:  # windows
        wb = load_workbook('自動查最低價.xlsm')

    ws = wb.active
    '''讀取 Excel 檔取得商品名單'''
    return [cell.value for cell in list(ws.columns)[1][2:] if not cell.value == None]

def is_same_prod(prod, found, color, threeC):
    '''判斷兩個品名是否為同一商品。\n
    ＊不去特別分辨 試用品 或 非試用品'''
    logging.info('-----------------------\n比對：\n%s\n%s', prod, found)

    # 兩品名皆轉換為小寫： 排除大小寫差異
    prod = removeComment(prod).lower()
    found = removeComment(found).lower()

    ___exist = False  # default. It will be True if there is '___' in prod
    if '___' in prod:
        ___exist = True

    # 移除標點符號
    prod = removePunc(prod)
    found = removePunc(found)

    # 遇到需要比較相似度的地方，我會用 re.sub('___','',prod) 來移除 prod 中的 '___'
    if SequenceMatcher(None, re.sub('___', '', prod), found).quick_ratio() == 1:
        logging.info('OOO: %s \n 和\n%s 吻合！', re.sub('___', '', prod), found)
        return True

    elif SequenceMatcher(None, re.sub('___', '', prod), found).quick_ratio() > 0.76:
        logging.info('找到相似度超過 0.76 的商品')
        logging.info('第一階段商品名處理：小寫化、移除標點符號、移除宣傳語')
        logging.info('prod=%s', prod)
        logging.info('found=%s', found)

        # 檢查「指定規格」是否一致
        for word in color:  # set color 的內容由使用者指定 (boom_data.xlsx)
            if word in prod and word not in found:
                logging.info('XXX: 指定檢查的規格不一致')
                print('指定檢查的規格不一致\n')
                return False

        # 如果商品是 3C ，用另外的特殊函式去判斷是否為相同商品 (目標是所有 3C 都在這邊處理)
        for brand in threeC:
            if brand.lower() in prod:
                logging.info('商品和 3C 有關，套用 3C 專用規格比較法')
                return is_same_specifi(prod, found, ___exist = ___exist)

        # 特別處理：split 後前兩個字串都不是中文 （防：英文字太多，會使中文字串的兩三字間的差異（規格）被忽略）
        # 我承認這裡怪怪的，應該可以改得更有效率
        if not is_chinese(prod.split()[0]) and len(prod.split()) > 1 and not is_chinese(prod.split()[1]):
            if SequenceMatcher(None, removeNoChinese(re.sub('___', '', prod)), removeNoChinese(found)).quick_ratio() > 0.75:
                s1 = ''
                s2 = ''

                # 抽掉它們不是中文的部分，重新比較
                for c in removeNoChinese(re.sub('___', '', prod)).split():
                    s1 += c

                for c in removeNoChinese(found).split():
                    s2 += c

                # 若是只拿中文的部分去互相比較相似度依然高，則再比對規格是否相符
                if SequenceMatcher(None, s1, s2).quick_ratio() > 0.75:
                    logging.info('pass: 只拿中文的部分去互相比較相似度依然高')
                    logging.info('prod_chineseOnly=%s', s1)
                    logging.info('found_chineseOnly=%s', s2)

                    pass

                else:
                    logging.info('XXX: 只拿中文的部分去互相比較，相似度不足 0.75')
                    return False

            else:
                logging.info('XXX: 相似度不足 0.75')
                return False

    # 預防：品名實際一樣，但因為有其中一方有英文版的品牌名，一方沒有，導致相似度過低
    # 對應：拿掉品名中的所有英文，重比相似度
    # 20201123 筆：Error 百出，3.4 版中直接拿掉了這一塊
    # 20201112 筆: 這個部分常常誤殺第一段子字串並非英文的正確搜尋結果，要修改（參照隔壁的 jupyter note)
    elif SequenceMatcher(None, re.sub('___', '', prod), found).quick_ratio() > 0.5 and (
            (len(prod.split()) > 1 and prod.split()[0].isalpha() and not is_chinese(prod.split()[0])) or (
            len(found.split()) > 1 and found.split()[0].isalpha() and not is_chinese(found.split()[0]))):
        logging.info('其中一方品名.split() 的第一段子字串不是中文，拿掉重新比較')
        logging.info('prod=%s', prod)
        logging.info('found=%s', found)
        logging.info('相似度大於 0.5 ，取走品名開頭非中文字串，再重新比較一次')

        if not is_chinese(prod.split()[0]):
            return is_same_prod(prod[prod.index(' ') + 1:], found, color)

        else:
            return is_same_prod(prod, found[found.index(' ') + 1:], color)


    # 有的時候查到的相符品名的末尾會被賣家硬塞很多關鍵字，無法通過相似度測驗
    elif len(found.split()) > 3 * len(prod.split()):
        logging.info('查到品名的末尾被賣家硬塞了很多關鍵字，重新檢查')
        same_count = 0
        for word in found.split():
            if word in prod:
                same_count += 1

        # 如果 found 的子字串（以空格分隔）有超過 3 串都有出現在 prod 的話，接著檢查規格
        if same_count > 3:
            logging.info('pass: 子字串檢查通過')
            pass

        else:
            logging.info('XXX: same_count=%d', same_count)
            return False

    # 相似度太低，排除
    else:
        logging.info('XXX: 相似度太低，排除')
        return False

    # Case1: 沒有要用'___'來當作隨意數字標記
    if not ___exist:  # 未完成
        #  ---檢查品名中出現的數字和數字順序(規格)是否一致---
        # 數字：指1,2,3,...。不包含one或一二三這種。
        numIn = ''
        numFoun = ''

        for n in [num if num.isdigit() else ' ' for num in prod]:
            numIn += n

        for n in [num if num.isdigit() else ' ' for num in found]:
            numFoun += n

        # 如果品名中出現的 數字 與 數字順序 不一致，則判斷為不同商品
        if numIn.split() == numFoun.split():
            # 不過濾試用品
            logging.info('OOO: 規格相符\nnumIn=%s\nnumFoun=%s', numIn, numFoun)
            return True

        else:
            logging.info('numIn=%s', numIn)
            logging.info('numFoun=%s', numFoun)

            print('查到：', found)
            print('規格不符\n')

            logging.info('XXX: 品名中出現的 數字 與 數字順序 不一致，判斷為不同商品')
            return False

    # Case2: 有要用'___'來當作隨意數字標記
    else:
        #  ---檢查品名中出現的數字和數字順序(規格)是否一致: 製作正則表達式---
        # 數字：指1,2,3,...。不包含one或一二三這種。
        numIn = ''
        numFoun = ''

        for i in range(len(prod)):
            if prod[i].isdigit():
                numIn += prod[i]

            elif i < len(prod) - 2 and prod[i] == '_' and prod[i + 1] == '_' and prod[i + 2]:
                numIn += '\d*'

        for n in found:
            if n.isdigit():
                numFoun += n

        # 如果正則表達式檢查通過 (i.e., re.search() 回傳非 None 值），則回傳 True
        if re.search(numIn, numFoun):
            logging.info('OOO:\n搜尋：%s\n查到：%s\n正則表達式檢查通過！', prod, found)
            return True

        else:
            logging.info('XXX: 正則表達式檢查未通過\nnumIn=%s\nnumFoun=%s', numIn, numFoun)
            return False

def removeComment(astr):
    '''用正則表達式移除對搜尋沒太大幫助的宣傳語、註記'''

    # unicode編碼參考： https://zh.wikipedia.org/wiki/Unicode%E5%AD%97%E7%AC%A6%E5%88%97%E8%A1%A8#%E5%9F%BA%E6%9C%AC%E6%8B%89%E4%B8%81%E5%AD%97%E6%AF%8D
    newstr_list = re.sub(
        r'[\[【(「（][^（「\[【(]*(新品上市|任選|效期|專案|與.+相容|免運|折後|限定|獨家\d+折|福利品|現折|限時|安裝|適用|點數[加倍]*回饋|[缺出現司櫃]貨|結帳|促銷)[^\]【）(」]*[\]】）)」]|[(]([^/ ]+/ *){1,}[^/]+[)]|效期[\W]*\d+[./]\d+[./]*\d*|\d(選|色擇)\d|.(折後.+元.|[一二兩三四五六七八九十]+色|([黑紅藍綠橙黃紫黑白金銀]/)+.|\w選\w色|只要.+起)|[^\u0020-\u0204\u4e00-\u9fa5]|[缺出現司櫃]貨[中]*|[^ ]*安裝[^ ]*|下架|[^ ]*配送',
        ' ', astr, 6).split()

    newstr = ''
    # 去除頭尾空白字元
    for word in newstr_list:
        newstr += word + ' '

    return newstr[:-1]

    # 去除頭尾空白字元
    for word in newstr_list:
        newstr += word + ' '

    return newstr[:-1]

def is_same_specifi(prod1, prod2, ___exist = False):
    '''判斷規格是否一致，和 is_same_prod 中的判斷法不同的是，數字出現的順序不必要一樣\n
    ___exist: flag; True: 使用者 有 使用'___'替代任意數字，反之為 False\n
    '''

    # iPad 篩選區： # 未完成
    if 'ipad' in prod1:
        # 判斷點1: iPad+數字 (ex: iPad7)
        if re.search('ipad *\d+', prod1) and re.search('ipad *\d+', prod1).group(0) in prod2:
            pass

        # 判斷點2： 年份 (ex: 2019)
        elif re.search('20\d\d', prod1) and re.search('20\d\d', prod1).group(0) in prod2:
            pass

        # 若有發現其他判斷點，則以 elif 繼續新增

        # 不符合此二判斷點的，都不是相同商品
        else:
            return False

    # 跳轉到這裡，商品英文已全切成小寫且移除非必要標點符號
    # 將商品中的中文全都抽取掉
    prod1 = removeChinese(prod1).split()
    prod2 = removeChinese(prod2).split()

    prod1_remove_allEng = [w for w in prod1 if not w.isalpha()]
    prod2_remove_allEng = [w for w in prod2 if not w.isalpha()]

    # print(prod1_remove_allEng)
    # print(prod2_remove_allEng)

    return SequenceMatcher(None, prod1_remove_allEng, prod2_remove_allEng).quick_ratio() == 1

def is_chinese(ustr):
    '''判斷一個字串是否全為中文字。 ＊只接受字串，不要亂丟int或其他什麼進來＊'''
    # 註：此處的中文字理論上包含所有繁體與簡體字
    # 只要字串包含以下任何一種，就會回傳 False ：全形標點符號、空白字元、英文、數字
    for uchar in ustr:
        if not (uchar >= u'\u4e00' and uchar <= u'\u9fa5'):
            return False

    return True

def removeNoChinese(astr):
    '''只保留原字串的中文與空白字元'''
    new_str = ''
    for c in astr:
        if is_chinese(c) or c in ' ':
            new_str += c
        else:
            new_str += ' '

    return new_str

def removeChinese(astr):
    '''移除字串中所有中文字元'''
    new_str = ''
    for c in astr:
        if not is_chinese(c):
            new_str += c

    return new_str

def removePunc(astr):
    '''移除對商品名稱而言多餘（拿掉也不應該影響搜尋）的標點符號'''
    result = ''
    strl = list(astr)

    # 不用拿掉 + * \' .
    while set(strl).intersection(
            {'》', '《', '「', '」', '【', '】', '!', '"', '#', '$', '%', '&', "'", '(', ')', ',', '-', '/', ':', ';', '<', '=', '>', '?', '@', '[', '\\',
             ']', '^', '`', '{', '|', '}', '~'}) != set():
        i = 0
        for c in strl:
            if c in '《》!"#$%,;<=>?@\\^`|':  # 把這些符號移除
                strl.remove(c)

            elif c in '&()[]{}:-~/「」【】\'':  # 把這些符號替換成空白
                strl.insert(i, ' ')
                strl.remove(c)

            i += 1

    # 把新字串拼回來
    for c in strl:
        result += c

    return result

def only_leave_alnumblank(astr):
    '''把一個字串中的數字、英文、空白符號以外的字元都移除'''
    newstr = ''

    for c in astr:
        if c.isalnum() or c.isspace():
            newstr = newstr + c

    return newstr

def removeComma_and_toInt(str_list):
    '''移除一個由數字字串組成的 list/ str 中的每個字串中的標點符號(',','$','~','～')，並轉換字串為 int type'''
    # Last Chnaged: 2020.9.27

    # --- 處理傳入的變數並不是 list 類型的狀況 ---
    if type(str_list) == str:  # str
        astr = str_list

        while ',' in astr:
            astr = astr[:astr.find(',')] + astr[astr.find(',') + 1:]

        while '$' in astr:
            astr = astr[:astr.find('$')] + astr[astr.find('$') + 1:]

        while '~' in astr:
            astr = astr[:astr.find('~')] + astr[astr.find('~') + 1:]

        while '～' in astr:
            astr = astr[:astr.find('～')]  # 取左邊的

        while '元' in astr:
            astr = astr[:astr.find('元')]

        return int(astr)

    elif type(str_list) == int:  # int
        return str_list  # 你幹嘛送不需要變換的東西進來啦
    # ---------------------------------------

    # --- 剩下的情形，預期：變數類型是 list ---
    if str_list == None:
        return None

    newlist = []
    if str_list == []:
        pass

    else:
        for astr in str_list:  # 2020.9.1 更新
            if type(astr) == int:
                newlist.append(astr)
                continue

            while ',' in astr:
                astr = astr[:astr.find(',')] + astr[astr.find(',') + 1:]

            while '$' in astr:
                astr = astr[:astr.find('$')] + astr[astr.find('$') + 1:]

            while '~' in astr:
                astr = astr[:astr.find('~')] + astr[astr.find('~') + 1:]

            while '～' in astr:
                astr = astr[:astr.find('～')]  # 取左邊的

            while '元' in astr:
                astr = astr[:astr.find('元')]

            newlist.append(int(astr))

    return newlist

# ========= 賣場可賣量抓取/ 價錢比對(isUrlAvailiable) ===========
def getAmount_PC(urls):
    '''使用 selenium 抓取 PChome 商品頁面的最大可賣量\n
    urls 是裝有網址的字串 list，回傳一個正整數的 list '''
    logging.info('Start getAmount_PC(urls)')

    # 無頭模式 報錯: 原始碼：loading 網頁載入中
    # 不知怎樣可解決此問題
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')  # 解決 webdriver automation extension 打不開的問題
    chrome_options.add_argument('user-agent=' + UserAgents[random.randint(0, len(UserAgents) - 1)])  # 解決抓到的原始碼是 loading 網頁載入中 的問題

    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')

    # driver = webdriver.Chrome(os.path.split(os.path.realpath(__file__))[0] + r'\chromedriver.exe', options = chrome_options)
    driver = webdriver.Chrome('chromedriver.exe', options = chrome_options)
    # wait = WebDriverWait(driver, 20)

    amounts = []
    j = 1
    for url in urls:
        amount = '-'

        if url == '-':
            amounts.append('-')
            continue

        driver.get(url)
        driver.implicitly_wait(30)
        # wait.until(EC.presence_of_element_located((By.CLASS_NAME,'Qty'))) 去掉這行好像就不會有視窗載不出來的問題
        soup = bs(driver.page_source, 'html.parser')
        # try:

        # Case1: 網頁顯示：商品售完補貨中
        if soup.find('li', id = 'ButtonContainer') is not None and '售完，補貨中！' in soup.find('li', id = 'ButtonContainer').text:
            amount = '售完補貨中'

        # Case2: 網頁顯示：商品完售
        elif soup.find('li', id = 'ButtonContainer') is not None and '完售，請參考其他商品' in soup.find('li', id = 'ButtonContainer').text:
            amount = '完售'

        # Case3: 以上兩者都沒發生：預期商品頁面有顯示商品可賣量
        else:
            try:
                amount = int(soup.find('select', class_ = 'Qty').find_all('option')[-1].text)

            except AttributeError:
                print("soup.select('select'):", soup.select('select'))
                print("soup.select('button')", soup.select('button'))
                # print("soup.find('select', class_='Qty').find_all('option')[-1].text=",soup.find('select', class_='Qty').find_all('option')[
                # -1].text) # AttributeError: 'NoneType' object has no attribute 'find_all'
                amount = 'AttributeError'
                time.sleep(30)

        '''except: # 完售
            print('error')
            print()
            amount = 'error'''
        amounts.append(amount)

        print('可賣量查詢進度：{}/{}'.format(j, len(urls)))
        j += 1

    driver.quit()
    return amounts

def getAmount_momo(driver, wait, url):
    '''取得 momo 商品網址內的商品可賣量\n
    一次只讀入一個網址'''
    logging.info('Start getAmount_momo(driver, wait, url)')

    driver.get(url)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'CompareSel')))
    soup = bs(driver.page_source, 'html.parser')
    amount = int(soup.find('select', class_ = 'CompareSel', id = 'count').find_all('option')[-1].text)

    return amount

def getAmount_etmall(urls):
    '''使用 selenium 抓取 東森購物 商品頁面的最大可賣量\n
    urls 是裝有網址的字串 list，回傳一個正整數的 list 。'''
    logging.info('Start getAmount_etmall(urls)')

    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')  # 解決 webdriver automation extension 打不開的問題
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    # chrome_options.add_argument('user-agent='+UserAgents[random.randint(0,len(UserAgents)-1)])

    # ---- 開啟 chrome driver  -----
    if platform.system() == 'Windows':  # Windows
        # driver = webdriver.Chrome(os.path.split(os.path.realpath(__file__))[0] + r'\chromedriver.exe', options = chrome_options)
        driver = webdriver.Chrome('chromedriver.exe', options = chrome_options)

    elif platform.system() == 'Darwin':  # Mac OS
        driver = webdriver.Chrome('./chromedriver', options = chrome_options)
    # ---- chrome driver 準備完畢 ----

    amounts = []
    j = 1
    for url in urls:
        if url == '-':
            amounts.append('-')
            continue
        driver.get(url)
        driver.implicitly_wait(5)
        soup = bs(driver.page_source, 'html.parser')
        amount = 'No_Found'

        try:
            if soup.find('select', class_ = 't-quantitySelector n-form--control') is not None:
                amount = int(soup.find('select', class_ = 't-quantitySelector n-form--control').find_all('option')[-1].text)

            elif soup.find('select', class_ = 't-quantitySelector') is not None:
                amount = int(soup.find('select', class_ = 't-quantitySelector').find_all('option')[-1].text)

        except AttributeError:
            print('AttributeError')

            if soup.find('a', class_ = 'n-btn n-btn--disabled') is not None and '銷售一空' in soup.find('a', class_ = 'n-btn n-btn--disabled').text:
                amount = '售完'

            else:
                print('未知錯誤，印出原始頁面內容：')
                print(soup.text)
                time.sleep(20)
                amount = 'error'

        amounts.append(amount)
        print('進度：{}/{}'.format(j, len(urls)))
        j += 1

    driver.quit()
    return amounts

def getStatus_Ymall(urls):
    '''抓取商品們處於可買或已售完狀態。回傳 list'''
    logging.info('Start getStatus_Ymall(urls)')

    s = requests.session()
    status = []

    for url in urls:
        req = s.get(url, headers = headers)
        soup = bs(req.text, 'html.parser')
        if '售完補貨中' in soup.find('div', id = 'ypsa2ct-2015').text:
            status.append('售完')

        elif '立即購買' in soup.find('div', id = 'ypsa2ct-2015').text:
            status.append('可買')

        else:  # 我也很好奇這會是什麼狀況
            status.append('兩者皆非')
            print(soup.find('div', id = 'ypsa2ct-2015').text)

        time.sleep(random.randint(1, 10) * 0.1)

    return status

def isUrlAvailiable(url, shop, price):
    '''檢查在 FindPrice 上所查到的商品價格是否正確、店面是否存在'''
    logging.info('Start isUrlAvailiable(url, shop, price)')

    r = requests.get(url, headers = headers)
    r.encoding = 'utf-8'
    soup = bs(r.text, 'html.parser')

    if shop == '樂天市場':
        try:
            return removeComma_and_toInt(soup.select_one('#auto_show_prime_price > strong > span').text) == price

        except:
            print(url)
            if soup.select_one('#auto_show_prime_price > strong > span') == None:
                print('[樂天市場] 找不到價格欄\n')
            else:
                print('[樂天市場] 價錢和Find Price 所顯示的不符\n')

            return False

    elif shop == 'Yahoo奇摩超級商城':
        try:
            return soup.find('span', class_ = 'price').text[:-1] == str(price)
        except:
            print(url)
            if soup.find('span', class_ = 'price') == None:
                print('[Yahoo奇摩超級商城] 找不到價格欄\n')
            else:
                print('[Yahoo奇摩超級商城] 價錢和Find Price 所顯示的不符\n')
            return False

    elif shop == 'Yahoo奇摩購物中心':
        try:
            return removeComma_and_toInt(soup.select_one(
                '#isoredux-root > div > div.ProductItemPage__pageWrap___2CU8e > div > div:nth-child(1) > div.ProductItemPage__infoSection___3K0FH > div.ProductItemPage__rightInfoWrap___3FNQS > div > div.HeroInfo__heroInfo___1V1O8 > div > div.HeroInfo__leftWrap___3BJHV > div > div').text) == price
        except:
            print(url)
            if soup.select_one(
                    '#isoredux-root > div > div.ProductItemPage__pageWrap___2CU8e > div > div:nth-child(1) > div.ProductItemPage__infoSection___3K0FH > div.ProductItemPage__rightInfoWrap___3FNQS > div > div.HeroInfo__heroInfo___1V1O8 > div > div.HeroInfo__leftWrap___3BJHV > div > div') == None:
                print('[Yahoo奇摩購物中心] 找不到價格欄\n')
            else:
                print('[Yahoo奇摩購物中心] 價錢和Find Price 所顯示的不符\n')
            return False

    elif shop == 'myfone購物':
        try:
            return removeComma_and_toInt(soup.select_one(
                '#item-419 > div.wrapper > div.section-2 > div.prod-description > div.prod-price > span.prod-sell-price').text) == price

        except:
            print(url)
            if soup.select_one('#item-419 > div.wrapper > div.section-2 > div.prod-description > div.prod-price > span.prod-sell-price') == None:
                print('[myfone購物] 找不到價格欄\n')
            else:
                print('[myfone購物]價錢和Find Price 所顯示的不符\n')

    elif shop == 'momo購物網':
        pass

    elif shop == 'PChome 24h購物':
        try:
            return soup.select_one('#PriceTotal').text == str(price)

        except:
            print(url)
            if soup.select_one('#PriceTotal') == None:
                print('[PChome 24h購物]找不到價格欄\n')
            else:
                print('[PChome 24h購物]價錢和Find Price 所顯示的不符\n')

            return False

    elif shop == 'ETmall東森購物網':
        try:
            return removeComma_and_toInt(soup.select_one(
                '#productDetail > div:nth-child(2) > section > section > div:nth-child(3) > div.n-price__block > div.n-price__bottom > span.n-price__exlarge > span.n-price__num').text) == price
        except:
            print(url)
            if soup.select_one(
                    '#productDetail > div:nth-child(2) > section > section > div:nth-child(3) > div.n-price__block > div.n-price__bottom > span.n-price__exlarge > span.n-price__num') == None:
                print('[ETmall東森購物網] 找不到價格欄\n')
            else:
                print('[ETmall東森購物網] 價錢和Find Price 所顯示的不符\n')

            return False

    elif shop == '蝦皮商城':
        pass

    elif shop == 'Costco 好市多線上購物':
        pass
    elif shop == 'udn買東西':
        pass

    elif shop == 'friDay購物':
        try:
            return soup.select_one(
                '#E3 > div > div > div.prodinfo_area > span > div.bayPricing_area > div.attract_block > span.useCash > span.price_txt').text == str(
                price)
        except:
            print(url)
            if soup.select_one(
                    '#E3 > div > div > div.prodinfo_area > span > div.bayPricing_area > div.attract_block > span.useCash > span.price_txt') == None:
                print('[friDay購物] 找不到價格欄\n')
            else:
                print('[friDay購物] 價錢和Find Price 所顯示的不符\n')
            return False

    elif shop == 'PChome 商店街':
        pass
    elif shop == '家樂福線上購物網':
        pass
    elif shop == 'momo摩天商城':
        try:
            return removeComma_and_toInt(soup.select_one(
                '#goodsForm > div.prdInnerArea > div > div.prdrightwrap > div.prdleftArea > div.prdDetailedArea > dl > dd.sellingPrice > span').text) == price
        except:
            print(url)
            if soup.select_one(
                    '#goodsForm > div.prdInnerArea > div > div.prdrightwrap > div.prdleftArea > div.prdDetailedArea > dl > dd.sellingPrice > span') == None:
                print('[momo摩天商城] 找不到價格欄\n')
            else:
                print('[momo摩天商城] 價錢和Find Price 所顯示的不符\n')

    elif shop == '森森購物網':
        pass
    elif shop == '愛買線上購物':
        pass
    elif shop == '大樹健康購物網':
        pass
    elif shop == '小三美日平價美妝':
        pass
    elif shop == '松果購物':
        pass
    elif shop == '生活市集':
        pass
    elif shop == 'Jollybuy 有閑':
        pass
    elif shop == '熊媽媽買菜網':
        pass
    elif shop == '淘寶精選':
        pass

'''def getAmount_FP(urls, shops):
'''
