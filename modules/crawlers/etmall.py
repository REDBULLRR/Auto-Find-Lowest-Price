from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_etmall(prod_list,ws, color, threeC):
    '''抓取 東森購物 上的最低價'''
    logging.info('Crawl on etmall')

    et_url = 'https://www.etmall.com.tw/Search/Get'   # 取得 json 用（從封包 GET 中取得)
    et_mainurl = 'https://www.etmall.com.tw'          # 拼接出商品頁面網址用

    # 印分隔線
    print('*'*15+' 東森購物網 '+'*'*15)
    result_for_write = []
    s = requests.Session()

    j=0
    try:
        for prod in prod_list:
            j+=1
            ws['a1'].value = ['東森：{0}/{1}'.format(j,len(prod_list))]
            #result[prod]['etmall'] = {}
            names = []
            urls = []
            image_urls = []
            prices = []

            page = 0
            payload = {
                'keyword':re.sub('___',' ',prod),
                'model[cateName]':'全站',
                'model[page]':0,
                'model[storeID]':'',
                'model[cateID]':-1,
                'model[filterType]':'',
                'model[sortType]':'',
                'model[moneyMaximum]':'',
                'model[moneyMinimum]':'',
                'model[pageSize]':'48',
                'model[SearchKeyword]':'',
                'model[fn]':'',
                'model[fa]':'',
                'model[token]':'',
                'model[bucketID]':1,
                'page':page
            }
            headers={
                "User-Agent":UserAgents[(random.randint(0,len(UserAgents)-1))]
                }
            req = s.post(et_url, data=payload, headers=headers) # 網頁會以 json 格式回傳表單搜尋結果

            if req.ok:
                print('搜尋商品：{}\n'.format(re.sub('___',' ',prod)))
            else:
                logging.warning('連線失敗')
                raise ConnectionError('連線失敗，狀態碼: {}，請檢查網路連線。\n'.format(req.status_code))

            req.encoding = 'utf-8'
            data = json.loads(req.text)

            # 無搜尋結果 --> 接著搜下一個商品  *改進計畫：考慮去過濾「相似商品」
            if data['searchResult']['isFuzzy']:
                logging.info("data['searchResult']['isFuzzy'] == False, 無搜尋結果")
                print('商品：{}\n無搜尋結果'.format(re.sub('___',' ',prod)))
                print('='*35)
                result_for_write.append(['-','-','-','-'])
                continue

            # 有搜尋結果 --> 開始過濾
            else:
                print('一共有{}頁搜尋結果. . .'.format(data['searchResult']['totalPages']))
                print('一共有{}筆資料. . .\n'.format(data['searchResult']['totalProducts']))

                # - - - - 讀取每一頁的結果 - - - - -
                for i in range(data['searchResult']['totalPages']):
                    logging.info('第 %d 頁'%(page+1))
                    print('- 第 {} 頁 -'.format(page+1))

                    # - - - 存取每頁的品名、賣價、商品連結、商品圖片網址 - - -
                    for item_dict in data['searchResult']['products']:
                        name = item_dict['title']                   # 品名
                        link = et_mainurl+item_dict['pageLink']     # 商品頁面網址
                        imageUrl = 'https'+item_dict['imageUrl']    # 商品圖片網址
                        price = item_dict['finalPrice']             # 商品賣價

                        #print(is_same_prod(prod, name))

                        # - - - 把和給定品名相符的品名資料存入 list - - -
                        if is_same_prod(prod, name, color,threeC):
                            logging.info('查到：%s'%name)
                            print('查到:{}'.format(name))
                            names.append(name)
                            urls.append(link)
                            image_urls.append(imageUrl)
                            prices.append(int(price))

                    page = i+1
                    payload = {
                        'keyword':re.sub('___',' ',prod),
                        'model[cateName]':'全站',
                        'model[page]':0,
                        'model[storeID]':'',
                        'model[cateID]':-1,
                        'model[filterType]':'',
                        'model[sortType]':'',
                        'model[moneyMaximum]':'',
                        'model[moneyMinimum]':'',
                        'model[pageSize]':'48',
                        'model[SearchKeyword]':'',
                        'model[fn]':'',
                        'model[fa]':'',
                        'model[token]':'',
                        'model[bucketID]':1,
                        'page':page
                    }
                    req = s.post(et_url, data=payload, headers=headers)
                    req.encoding = 'utf-8'
                    data = json.loads(req.text)

            if names == []:
                logging.info('name == [], 無相符搜尋結果')
                print('無相符搜尋結果')
                print('='*35)
                result_for_write.append(['-','-','-','-'])
                continue

            else:
                print('共有{}項相符搜尋結果'.format(len(names)))
                try:
                    result_for_write.append([min(prices),names[prices.index(min(prices))], '正在查詢', urls[prices.index(min(prices))]])

                except:
                    logging.exception('Error occour when crawling PChome!')
                    logging.info('names=%s',names)
                    logging.info('prices=%s',prices)
                    logging.info('urls=%s',urls)
                    logging.info('result_for_write=%a',result_for_write)

                ##result[prod]['etmall']['imgUrl'] = image_urls[prices.index(min(prices))]
                ##result[prod]['etmall']['prodUrl'] = urls[prices.index(min(prices))]
                print('='*35)

    except:
        logging.exception('Error occour when crawling PChome!')

    finally:
        # 將資料寫入儲存格
        ws['q3'].value = result_for_write
        logging.info('東森：可賣量查詢中')
        ws['a1'].value = '東森：可賣量查詢中'

        urls = [alist[-1] for alist in result_for_write]
        amounts = getAmount_etmall(urls)
        logging.info('amounts=%d',amounts)
        amounts_to_write = [[num] for num in amounts]
        ws['s3'].value = amounts_to_write

        logging.info('東森購物網抓取完畢！\n')
        print('東森購物網抓取完畢！\n')
