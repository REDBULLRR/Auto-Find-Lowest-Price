from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_momo(prod_list, ws, color, threeC):
    '''用 selemium 上 momo 抓取商品最低價'''
    logging.info('Crawl on momo')

    momo_url = 'https://www.momoshop.com.tw/main/Main.jsp'  # 搜尋頁面
    momo_mainurl = 'https://www.momoshop.com.tw/'  # 拼接商品網址用

    print('*' * 15 + ' MOMO ' + '*' * 15)
    ws['a1'].value = '連線到 MOMO. . .'
    result_for_write = []

    # - - - 啟動無頭模式 - - -
    chrome_options = Options()
    chrome_options.add_argument('--no-sandbox')  # 看這行能否解決 webdriver automation extension 打不開的問題
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    # chrome_options.add_argument('user-agent='+UserAgents[random.randint(0,len(UserAgents)-1)])
    # 加了這行的話會在  searchBox = driver.find_element_by_name('keyword') 這行報錯：   can't find this element
    # - - - 打開瀏覽器 - - -
    pass

    # ---- 開啟 chrome driver  -----
    if platform.system() == 'Windows':  # Windows
        # driver = webdriver.Chrome(os.path.split(os.path.realpath(__file__))[0]+r'\chromedriver.exe',options = chrome_options)
        driver = webdriver.Chrome('chromedriver.exe', options = chrome_options)

    elif platform.system() == 'Darwin':  # Mac OS
        driver = webdriver.Chrome('./chromedriver', options = chrome_options)
    # ---- chrome driver 準備完畢 ----

    wait = WebDriverWait(driver, 5)

    try:
        # - - - 連線到 momo 前台主頁網址 - - -
        driver.get(momo_url)

        j = 1
        # - - - 抓取清單內商品 - - -
        for prod in prod_list:
            ws['a1'].value = 'MOMO：{0}/{1}'.format(j, len(prod_list))
            print('搜尋商品：{} \n'.format(re.sub('___', ' ', prod)))
            j += 1
            # result[prod]['momo'] = {}

            # - - - 找到搜尋欄，輸入商品名，按下 Enter - - -
            searchBox = driver.find_element_by_name('keyword')
            searchBox.clear()
            searchBox.send_keys(re.sub('___', ' ', prod), Keys.RETURN)

            # - - - 等搜尋結果載出來 - - -
            try:
                # pattern1: 有相符搜尋結果/ 無相符搜尋結果，給相似商品
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "totalTxt")))
            except:
                # pattern2: 連相似商品都沒有
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'newSearchFailsArea')))
                print('無搜尋結果')
                print('=' * 35)
                result_for_write.append(['-', '-', '-', '-'])
                continue

            # - - 抓取「綜合排序」的搜尋結果 - -
            # - - - - - - 解析原始碼 - - - - - - -
            root = bs(driver.page_source, 'html.parser')

            # - - - - 找到所有價格、品名、庫存- - - - - -
            listArea = root.find('div', class_ = 'listArea')

            # print(listArea)
            prices_found = [x.text[1:] for x in listArea.find_all('span', class_ = 'price')]
            names_found = [x.text for x in listArea.find_all('h3', class_ = 'prdName')]
            urls_found = [momo_mainurl + ele['href'] for ele in listArea.select('a.goodsUrl')]
            img_urls_found = [ele['src'] for ele in listArea.select('img.prdImg')]

            names = []
            prices = []
            urls = []
            img_urls = []

            for i in range(len(names_found)):
                if is_same_prod(prod, names_found[i], color, threeC):
                    names.append(names_found[i])
                    prices.append(prices_found[i])
                    urls.append(urls_found[i])
                    img_urls.append(img_urls_found[i])

            if names == []:
                print('此商品在 MOMO 無相符搜尋結果\n')
                result_for_write.append(['-', '-', '-', '-'])
                print('=' * 35)

            else:
                print('抓到相符商品數：', len(names), '\n')

                # 將資料存入
                # 庫存待補
                url = urls[prices.index(min(prices))]
                result_for_write.append([min(prices), names[prices.index(min(prices))], getAmount_momo(driver, wait, url), urls[prices.index(min(prices))]])

                # result[prod]['momo']['lowest_price'] = min(prices)
                # result[prod]['momo']['foundName'] = names[prices.index(min(prices))]
                # result[prod]['momo']['imgUrl'] = img_urls[prices.index(min(prices))]
                # result[prod]['momo']['prodUrl'] = urls[prices.index(min(prices))]
                print('=' * 35)

            # 等一等，當個有禮貌的爬蟲（免得被鎖IP）
            time.sleep(random.randint(1, 10) * 0.05)

        print('momo 抓取完畢！\n')

    except:
        logging.exception('Error occour when crawling MOMO!')
        ws['i3'].value = '未知錯誤'

    finally:
        ws['i3'].value = result_for_write
        # - - - - 結束瀏覽器，釋放記憶體空間 - - - -
        driver.quit()
