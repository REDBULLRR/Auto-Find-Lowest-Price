from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_Ybuy(prod_list, ws, color, threeC):
    '''抓取 Y購 上的最低價'''
    logging.info('Crawl on Ybuy')

    # 問題：搜尋不到商品
    # 策略：拆為關鍵字搜尋

    Yb_url = 'https://tw.buy.yahoo.com/search/product?'
    Yb_official_url = 'https://tw.buy.yahoo.com/?sub=283'

    # 印分隔線
    print('*'*15+' Yahoo!購物中心 '+'*'*15)
    result_for_write = []
    s = requests.session()

    j=0
    try:
        for prod in prod_list:
            j+=1
            ws['a1'].value = 'Y購：{}/{}'.format(j,len(prod_list))
            logging.info('搜尋商品：%s'%re.sub('___',' ',prod))
            print('搜尋商品：{}\n'.format(re.sub('___',' ',prod)))
            #result[prod]['Ybuy'] = {}
            names = []
            urls = []
            prices = []

            req = s.get(Yb_url, params={'p':re.sub('___',' ',prod),'sort':'price'}, headers=headers)
            req.encoding = 'utf-8'
            # test
            #ws['a1'].value = req.url

            soup = bs(req.text ,'html.parser')

            if soup.find('ul',class_='gridList')== None:
                # Y拍賣上有相似商品：改去Y拍上面搜尋
                if soup.select_one('#isoredux-root > div.page.shopping > div > div > div > div > div:nth-child(3) > div > span.NoResult_title_2bt6D'):
                    req = s.get()

                else:
                    logging.info('Ybuy: 無任何搜尋結果')
                    print('無任何搜尋結果')
                    print('='*35)
                    result_for_write.append(['-','-','-'])
                    continue

            found_names = [ele.text for ele in soup.select('span.BaseGridItem__title___2HWui')]  # 爬取品名

            found_prices = [span.em.text for span in soup.select('span.BaseGridItem__itemInfo___3E5Bx')]# 爬取價格

            found_urls = [child.a['href'] for child in soup.find_all('ul',class_='gridList')[-1].children if child.a!=None and child.a['href']!=Yb_official_url] # 爬取商品商店網址

            # test
            #print('len(found_names):',len(found_names)) #5
            #print('len(found_prices):',len(found_prices))#4
            #time.sleep(3)


            '''這段不能用！這個只（？）抓得到旁邊有「劃掉的價錢」的價格搜尋結果'''
            #if len(found_prices) != len(found_names): # 有時候要用下面這行才抓得到價錢
            #    found_prices = [span.em.text for span in soup.select('span.BaseGridItem__price___31jkj')]
            #    print('呱')
            #    print('len(found_prices):',len(found_prices))#1
                #print(soup.find_all('ul',class_='gridList')[-1].text)


            # 檢查抓到的商品價格與網址是否為空
            if not found_prices or not found_urls:
                logging.warning('Search XXX: 抓到的商品價格或商品網址為空。目標網站結構可能有變，原始碼需要修改')
                logging.warning('抓到商品價格：%a',found_prices)
                logging.warning('抓到商品網址：%a',found_urls)
                print('目標網頁結構可能有變，沒有抓到目標網址或價格')
                print('跳過 Y城 的搜尋')
                break

            # 檢查爬到的「商品名筆數」是否和「商品價格筆數」一致
            if len(found_names)!=len(found_prices):
                logging.warning('從原始碼中抓到的品名數和價格數不同')
                logging.info('prod:%s',prod)
                logging.info('found_prices:%a',found_prices)
                logging.info('found_names:%a',found_names)
                logging.info('found_urls:%a',found_urls)
                logging.info("soup.select('em.BaseGridItem__price___31jkj')=%a",soup.select('em.BaseGridItem__price___31jkj'))

            try:
                # 過濾並存取相符品名
                for i in range(len(found_names)):
                    if is_same_prod(prod, found_names[i], color,threeC):
                        logging.info('查到：%s'%found_names[i])
                        print('查到：',found_names[i])
                        names.append(found_names[i])
                        prices.append(found_prices[i])
                        urls.append(found_urls[i])

            except:
                logging.exception('Ybuy:過濾並存取相符品名時出現錯誤')
                logging.info('prod:%s',prod)
                logging.info('found_prices:%a',found_prices)
                logging.info('found_names:%a',found_names)
                logging.info('found_urls:%a',found_urls)
                logging.info('i:%d',i)
                logging.info("soup.select('em.BaseGridItem__price___31jkj')=%a",soup.select('em.BaseGridItem__price___31jkj'))

            prices = removeComma_and_toInt(prices)

            if names == []:
                logging.info('無相符之搜尋結果')
                print('無相符之搜尋結果')
                print('='*35)
                result_for_write.append(['-','-','-'])
                continue

            logging.info('抓到相符商品數：%d',len(names))
            logging.info('抓到商品：%a',names)
            print('抓到相符商品數：',len(names))

            # 將資料存入
            result_for_write.append([min(prices),names[prices.index(min(prices))], urls[prices.index(min(prices))]])
            ##result[prod]['Ybuy']['lowest_price'] =
            ##result[prod]['Ybuy']['foundName'] =
            ##result[prod]['Ybuy']['prodUrl'] = urls[prices.index(min(prices))]
            print('='*35)

    except:
        logging.exception('Error occour when crawling Ybuy!')

    finally:
        # 將資料寫入儲存格
        ws['f3'].value = result_for_write

        print('Yahoo!購物中心 抓取完畢！\n')
