from modules.libraries import *
from modules.globals import *
from modules.utilities import *

def crawler_on_pchome(prod_list, ws, color, threeC):
    """搜尋並爬取 pchome 上的商品。使用 requests。"""
    logging.info('Crawl on PChome')

    print('*' * 15 + ' PChome ' + '*' * 15)
    result_for_write = []

    pc_url = 'https://ecshweb.pchome.com.tw/search/v3.3/all/results?'
    pc_mainurl = 'http://24h.pchome.com.tw/prod/'  # 取得商品網頁網址用
    pc_mainurl2 = 'https://d.ecimg.tw'  # 取得圖片網址用

    j = 0
    try:
        for prod in prod_list:
            j += 1
            ws['a1'].value = 'PChome：{0}/{1}'.format(j, len(prod_list))
            logging.info('搜尋商品：%s' % re.sub('___', ' ', prod))
            print('搜尋商品：{}\n'.format(re.sub('___', ' ', prod)))
            # result[prod]['PChome'] = {}
            names = []
            prices = []
            urls = []
            img_urls = []
            page = 0

            while page <= 10:
                page += 1
                try:
                    # --抓取網站回傳的 Json 格式資料 --
                    payload = {
                        'q'   : re.sub('___', ' ', prod),
                        'page': page,
                        'sort': 'rnk/dc'
                    }
                    resp = requests.get(pc_url, params = payload, headers = headers)
                    # print(resp.url)
                    resp.encoding = 'utf-8'
                    respp = resp.text
                    data = json.loads(respp)

                    if type(data) == list:
                        raise KeyError('這個頁面所抓到的 json 不是預期要抓的格式。')

                    for item in data['prods']:
                        if is_same_prod(prod, item['name'], color, threeC):
                            names.append(item['name'])
                            prices.append(item['price'])
                            urls.append(pc_mainurl + item['Id'])
                            img_urls.append(pc_mainurl2 + item['picB'])

                            print('查到：', item['name'])

                except KeyError:
                    print('爬到第{}頁為止'.format(page))
                    break

                except json.decoder.JSONDecodeError:  # 預期：連相似結果都沒有
                    print('JSONDecodeError')
                    # print(data)
                    break

            if names == []:
                print('此商品在 PChome 無相符搜尋結果\n')
                print('=' * 35)
                result_for_write.append(['-', '-', '-', '-'])
                continue

            else:
                print('相符商品筆數：{}\n'.format(len(names)))
                print('=' * 35)
                result_for_write.append([min(prices), names[prices.index(min(prices))], '稍後寫入', urls[prices.index(min(prices))]])
                # 將資料存入字典
                # result[prod]['PChome']['lowest_price'] = min(prices)
                # result[prod]['PChome']['foundName'] = names[prices.index(min(prices))]
                # result[prod]['PChome']['imgUrl'] = img_urls[prices.index(min(prices))]
                # result[prod]['PChome']['prodUrl'] = urls[prices.index(min(prices))]
                pass

            # --- 嘗試當個有禮貌的爬蟲 ---
            time.sleep(random.randint(1, 10) * 0.01)

    except:
        logging.exception('Error occour when crawling PChome!')

    finally:
        ws['m3'].value = result_for_write
        ws['a1'].value = ['正在查詢PChome商品可賣量']

        urls = [alist[-1] for alist in result_for_write]
        amounts = getAmount_PC(urls)
        print('amounts=', amounts)
        amounts_to_write = [[num] for num in amounts]
        ws['o3'].value = amounts_to_write

        print('PChome 抓取完畢！\n')
