# test
#prodList = ['LAVIN 浪凡 花漾公主女性淡香水 90ml TESTER','SAMSUNG三星Galaxy A71 5G 8G/128G 6.7吋智慧手機','健司 辻利抹茶奶茶沖泡飲 22g * 30包','Jo Malone 英國梨與小蒼蘭 香水 100ml','EBI ELIE SAAB 夢幻花嫁淡香精 TESTER 90ml','MONTBLANC 萬寶龍 海洋之心女性淡香水 30ml 試用品TESTER','豐力富 紐西蘭頂級純濃奶粉 2.6 公斤']

from modules import *

def main():
    os.chdir(os.path.dirname(__file__))
    logging.basicConfig(filename = 'shopCrawlers.log', filemode = 'a', level = logging.DEBUG)

    xw.Book.caller()
    prodList = getProdList()

    app = xw.apps.active
    wb = xw.books.active
    ws = xw.sheets.active

    logging.info('%s Mission Start!',time.asctime())
    color = set(ws['a14'].value.split(','))
    threeC = list(ws['a18'].value.split(','))

    # test
    print("color = ")
    print(color)
    print("3C = ")
    print(threeC)
    logging.info(json.dumps(list(color)))
    logging.info(json.dumps(threeC))
    logging.shutdown()  # flush log file

    crawler_on_Ymall(prodList,ws,color,threeC)
    crawler_on_Ybuy(prodList,ws,color,threeC)
    crawler_on_momo(prodList,ws,color,threeC)
    crawler_on_pchome(prodList,ws,color,threeC) # 時有 Bug 時沒有
    crawler_on_etmall(prodList,ws,color,threeC) 
    crawler_on_FP(prodList,ws,color,threeC) # 時好時壞

    ws['a1'].value = '抓好了'
    logging.info('%s Mission Complete!',time.asctime())
    os.system("pause")

def test():
    os.chdir(os.path.dirname(__file__))
    logging.basicConfig(filename = 'shopCrawlers.log', filemode = 'w', level = logging.DEBUG)
    logging.info('%s Mission Complete!', time.asctime())
    logging.debug('%s', __name__)
    # xw.serve()    # what is this? crashes Excel

if __name__ == '__main__':
    # executes when run from IDE, but not from Excel
    test()
