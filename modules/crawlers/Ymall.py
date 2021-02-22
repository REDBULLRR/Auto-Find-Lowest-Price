from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_Ymall(prod_list,ws, color, threeC):
    '''抓取 Y城 上的最低價'''
    logging.info('Crawl on Ymall')
    Ym_url = 'https://tw.search.mall.yahoo.com/search/mall/product?'

    # 印分隔線
    print('*'*15+' Yahoo!超級商城 '+'*'*15)
    result_for_write = []

    s = requests.session()

    j=0
    try:
        for prod in prod_list:
            j+=1
            ws['a1'].value = 'Y城：{0}/{1}'.format(j,len(prod_list))
            logging.info('搜尋商品：%s'%re.sub('___',' ',prod))
            print('搜尋商品：{}\n'.format(re.sub('___',' ',prod)))
            #result[prod]['Ymall'] = {}
            names = []
            urls = []
            prices = []

            req = s.get(Ym_url, params={'p':re.sub('___',' ',prod),'qt':'product','sort':'p'}, headers=headers)
            req.encoding = 'utf-8'
            logging.info('%s',req.url)

            soup = bs(req.text ,'html.parser')

            if  soup.find('ul',class_='gridList')== None:
                logging.info("soup.find('ul',class_='gridList')== None, 無任何搜尋結果")

                print('無任何搜尋結果')
                print('='*35)
                result_for_write.append(['-','-','-'])
                continue

            # ----- 抓取搜尋結果網頁上的品名、價格、網址 ------
            try:
                found_names = [x.string for x in soup.find('ul',class_='gridList').find_all('span', class_='BaseGridItem__title___2HWui')]

                found_prices = [x.string[1:] for x in soup.find('ul',class_='gridList').find_all('em')]

                found_urls = [li.a['href'] for li in soup.find('ul',class_='gridList').find_all('li',class_='BaseGridItem__grid___2wuJ7')]


            except TypeError: # type 轉換時的錯誤預防
                logging.warning('Ymall: TypeError')

                for li in soup.find('ul',class_='gridList').find_all('li'):
                    logging.info("搜尋到的商品列表：\nfor li in soup.find('ul',class_='gridList').find_all('li'):\nli.text=%s",li.text)

            if not found_prices or not found_urls:
                logging.warning('Search XXX: 抓到的商品價格或商品網址為空。目標網站結構可能有變，原始碼需要修改')
                logging.warning('抓到商品價格：%a',found_prices)
                logging.warning('抓到商品網址：%a',found_urls)
                print('目標網頁結構可能有變，沒有抓到目標網址或價格')
                print('跳過 Y城 的搜尋')
                break

            # 過濾並存取相符品名
            for i in range(len(found_names)):
                if is_same_prod(prod, found_names[i],color,threeC):
                    logging.info('查到：%s',found_names[i])
                    print('查到：',found_names[i])
                    names.append(found_names[i])
                    prices.append(found_prices[i])
                    urls.append(found_urls[i])

            prices = removeComma_and_toInt(prices)

            if names == []:
                logging.info('無相符之搜尋結果')
                print('無相符之搜尋結果')
                print('='*35)
                result_for_write.append(['-','-','-'])
                continue

            logging.info('抓到相符商品數：%d',len(names))
            logging.info('抓到商品：%a',names)
            print('抓到相符商品數：',len(names),'\n')
            # 將資料存入以待寫入
            result_for_write.append( [min(prices), names[prices.index(min(prices))], urls[prices.index(min(prices))]])

            ##result[prod]['Ymall']['prodUrl'] = urls[prices.index(min(prices))]
            print('='*35)

    except:
        logging.exception('Error in Ymall!')

    finally:
        # 將資料寫入儲存格
        ws['c3'].value = result_for_write
        logging.info('Yahoo!超級商城 抓取完畢！\n')
        print('Yahoo!超級商城 抓取完畢！\n')
