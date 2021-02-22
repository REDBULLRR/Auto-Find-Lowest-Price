from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_FP(prod_list, ws, color, threeC):
    '''抓取比價網站 Find Price 上的最低價'''
    logging.info('Crawl on FP')

    fp_mainurl = 'https://www.findprice.com.tw/g/'
    fp_mainurl2 = 'https://www.findprice.com.tw'

    # 印分隔線
    print('*' * 15 + ' Find Price ' + '*' * 15)
    result_for_write = []
    s = requests.session()

    j = 0
    try:
        for prod in prod_list:
            j += 1
            ws['a1'].value = 'FP：{0}/{1}'.format(j, len(prod_list))
            logging.info('搜尋商品：%s' % re.sub('___', ' ', prod))
            print('搜尋商品：{}\n'.format(re.sub('___', ' ', prod)))
            # result[prod]['FP'] = {}
            names = []
            urls = []
            prices = []
            shops = []

            url = fp_mainurl + quote(re.sub('___', ' ', prod))
            req = s.get(url, headers = headers)
            count = 1
            while not req.ok:
                logging.warning('連線 Find Price 失敗，請檢查網路狀態。程式接下來每隔三秒會重新嘗試連線')
                print('連線 Find Price 失敗，請檢查網路狀態。程式接下來每隔三秒會重新嘗試連線')
                time.sleep(3)

                req = s.get(url, headers = headers)
                count += 1
                if count > 5:
                    logging.error('連線失敗次數過多，自動跳出 Find Price')
                    raise TimeoutError('連線失敗次數過多，自動跳出 Find Price')

            req.encoding = 'utf-8'

            soup = bs(req.text, 'html.parser')

            if soup.find('div', id = 'GoodsGridDiv') == None or soup.find('div', id = 'GoodsGridDiv').a == None:
                # 有時候會出現：所有商品（包括人工搜時有搜尋結果的商品）都沒有搜尋結果的狀況
                # 原因應該是頁面抓不到這個 if 上面在判斷的 div tag
                logging.debug("沒有抓到 id='GoodsGridDiv' 的 div 標籤")
                logging.debug('req.json = %s' % req.json)
                logging.debug("[div.text for div in soup.find_all('div')] = %a" % [div.text for div in soup.find_all('div')])

                print('無任何搜尋結果')
                result_for_write.append(['-', '-', '-', '-'])

                print('=' * 35)
                continue

            # 麻煩的傢伙來了
            if soup.find('div', id = 'HotDiv').table != None:
                logging.info('FP: 麻煩的搜尋結果出現了，請前方人員進入戰鬥狀態')

                link = fp_mainurl2 + soup.find('div', id = 'HotDiv').a['href']
                reqq = s.get(link, headers = headers)
                reqq.encoding = 'utf-8'
                soup2 = bs(reqq.text, 'html.parser')

                found_names = [tr.find_all('td')[2].a.text for tr in soup2.find('div', id = 'GoodsGridDiv').find_all('tr')]

                found_prices = [tr.find_all('td')[1].text for tr in soup2.find('div', id = 'GoodsGridDiv').find_all('tr')]

                found_urls = [tr.find_all('td')[2].a['href'] for tr in soup2.find('div', id = 'GoodsGridDiv').find_all('tr')]

                found_shops = [tr.find_all('td')[2].img['title'] for tr in soup2.find('div', id = 'GoodsGridDiv').find_all('tr')]

                found_prices = removeComma_and_toInt(found_prices)

                logging.debug('唧')

                for i in range(len(found_names)):
                    if is_same_prod(prod, found_names[i], color, threeC):
                        logging.info('查到：%s' % found_names[i])
                        print('查到：', found_names[i])
                        names.append(found_names[i])
                        prices.append(found_prices[i])
                        urls.append(fp_mainurl2 + found_urls[i])
                        shops.append(found_shops[i])

                        logging.debug('呱')

            # 一般來說會直接用這邊的程式碼
            found_names = [tr.find_all('td')[1].a.text for tr in soup.find('div', id = 'GoodsGridDiv').find_all('tr')]

            found_prices = [tr.find_all('td')[1].span.text for tr in soup.find('div', id = 'GoodsGridDiv').find_all('tr')]

            found_urls = [tr.find_all('td')[1].a['href'] for tr in soup.find('div', id = 'GoodsGridDiv').find_all('tr')]

            found_shops = [tr.find_all('td')[1].img['title'] for tr in soup.find('div', id = 'GoodsGridDiv').find_all('tr')]

            try:
                for i in range(len(found_names)):
                    if is_same_prod(prod, found_names[i], color, threeC):
                        logging.info('查到：%s' % found_names[i])
                        print('查到：', found_names[i])
                        names.append(found_names[i])
                        prices.append(found_prices[i])
                        urls.append(fp_mainurl2 + found_urls[i])
                        shops.append(found_shops[i])

            except AttributeError:  # found_names[i] 是 list?
                logging.error('FP crawler: AttributeError')
                logging.info('names=', names)
                logging.info('prices=', prices)
                logging.info('urls=', urls)
                logging.info('shops=', shops)

            prices = removeComma_and_toInt(prices)

            if names == []:
                logging.info('無相符之搜尋結果')
                print('無相符之搜尋結果')
                print('=' * 35)

                result_for_write.append(['-', '-', '-', '-'])
                continue

            logging.info('抓到相符商品數：%d\n' % len(names))
            print('抓到相符商品數：', len(names), '\n')
            # 存下資料以待寫入
            result_for_write.append([min(prices), names[prices.index(min(prices))], shops[prices.index(min(prices))],
                                     urls[prices.index(min(prices))]])  # 不同平台的商店所對應的可賣量位置不同，待補。現階段只輸出網址
            print('=' * 35)

    except:
        logging.exception('Error occour when crawling Find Price')

    finally:
        # 將資料寫入儲存格
        ws['u3'].value = result_for_write
        logging.info('Find Price 抓取完畢！\n')
        print('Find Price 抓取完畢！\n')
