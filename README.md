# 自動查最低價4.0 筆記

## [【品名的搜尋關鍵字抽取函式】試行錯誤紀錄](https://hackmd.io/IIqZjuWqQv2hHS_BisxLsA)

## Project Setup
### xlwings
* `pip install xlwings`
    * had to reinstall python to get pip working...
    * if using "xlwings ......" in shell gives "failed to create process" error, reinstall python on a path *with no spaces*
* Q: `pip install --upgrade pip` breaks pip!
    * run `python -m ensurepip` to reinstall pip
    * use `python -m pip install --upgrade pip` as was indicated by the warning message
* Download xlwings.xlam
    * From https://github.com/xlwings/xlwings/releases
    * Run `xlwings addin install`, and if that fails...
    * Move `xlwings.xlam` to `C:\Users\{username}\AppData\Roaming\Microsoft\Excel\XLSTART` to enable xlwings tab in Excel
* Follow these instructions:
    * https://www.automateexcel.com/vba/install-add-in
* `xlwings quickstart MyTest` to create a project directory MyTest with a sample Excel and Python file
* To show console window when executing Python from Excel:
    * `%HOMEPATH%\.xlwings\xlwings.conf` > Add line `"SHOW CONSOLE","True"` > Restart Excel
### python libraries
* install as many as possible that are imported in `shopCrawlers.py`
    * some may fail to install, but it's ok
### chromedriver
* download chromedriver from https://chromedriver.chromium.org/ and add it to ~~PATH environmental variable~~ the same directory as the xlsm and py file
    * adding it to PATH still fails (see issues below for solution)


## Code Notes
* add "cookie" header to HTTP request to prevent IP lockdown
    * https://tlyu0419.github.io/2020/06/13/Crawler-PChome/
```python=1
headers = {
    'cookie': 'ECC=GoogleBot',
    # - - 使用隨機的 User-Agent - - - 
    "User-Agent":UserAgents[(random.randint(0,len(UserAgents)-1))]
}
```
* **Idea:** add an Excel sheet for each site to show all found results sorted by price, in case the script finds an incorrect item with lowest price
* shopCrawlers:
    * when run from IDE, `__name__ == __main__`
    * when run from Excel, `__name__ == 'shopCrawlers'`
* Logging levels
    * https://docs.python.org/3/library/logging.html#logging-levels
    * CRITICAL > ERROR > WARNING > INFO > DEBUG > NOTSET
* 難怪log file那麼肥，原來是把HTML都塞進去了XD
    * 用notepad開會慢爆，要用notepad++開
```python=1!
logging.info("搜尋到的商品列表：\nfor li in soup.find('ul',class_='gridList').find_all('li'):\nli.text=%s", li.text)
```
* **[Idea]:** If crawler fails (website structural change), display error message directly in Excel
    * users won't check the log file
* Comparisons can be chained in Python. wow.
    * https://stackoverflow.com/questions/26502775/simplify-chained-comparison
```python=1
# old
if not (uchar >= u'\u4e00' and uchar <= u'\u9fa5'):
# new
if not (u'\u4e00' <= uchar <= u'\u9fa5'):
```

## Code advice
* Good use of functions to separate functionality
    * Could put each function in a separate file (or "module" in Python language). All the code in one file makes IDE linting slow.
* Good commenting, but some of it (such as todo and problems to be solved) that does not directly concern the code can be put into a separate file
    * the import lines should be visible when at the top of file
* 自動查最低價.xlsm can be set as a parameter for flexibility (or user can set up a shortcut to the file if they wish to change its location)
    * what?
* Since load_workbook is only used twice
```python=1
# this code
from openpyxl import load_workbook
wb = load_workbook('自動查最低價.xlsm')
# can be changed into
```python=1
import openpyxl
wb = openpyxl.load_workbook('自動查最低價.xlsm')
```
for better readability
* `if not cell.value==None` :arrow_right: `if cell.value is not None`
    * more readable and *maybe* less error-prone
```python=1
return [cell.value for cell in list(ws.columns)[1][2:] if not cell.value==None]
```
* Put all the code in the Python file and just call the function from VBA
```python=1
# This VBA code
RunPython ("import logging; logging.basicConfig(filename= 'shopCrawlers.log', filemode='w', level=logging.DEBUG); import shopCrawlers; shopCrawlers.main()")
# should be simply
RunPython("import shopCrawlers; shopCrawlers.main();")
# everything else goes in shopCrawlers.py
```
* Better structure to collect all `<li>` elements and then extract name, price, url from each of them? (maybe)
    * But it works. I'm not touching it.
```python=1
gridList = soup.find('ul', class_ = 'gridList')
found_names = [x.string for x in gridList.find_all('span', class_ = 'BaseGridItem__title___2HWui')]
found_prices = [x.string[1:] for x in gridList.find_all('em', class_ = 'BaseGridItem__price___31jkj')]
found_urls = [li.a['href'] for li in gridList.find_all('li', class_ = 'BaseGridItem__grid___2wuJ7')]
```
* 是"儲存"不是"存取"ㄚㄚㄚ 存取是access的意思
```python=1
# 過濾並存取相符品名
for i in range(len(found_names)):
    if is_same_prod(prod, found_names[i], color, threeC):
	logging.info('查到：%s', found_names[i])
	print('查到：', found_names[i])
    names.append(found_names[i])
    prices.append(found_prices[i])
    urls.append(found_urls[i])
```
* To remove whitespaces:
```python=1
# old
newstr = ''
for word in newstr_list:
    newstr += word + ' '
return newstr[:-1]
# new
return newstr_list. # I forgot, will check next time
```

## Issues
1. **[Solved]** Cannot write log file
    * logging.basicConfig() by default writes log files to the Python installation directory, so do `os.chdir(os.path.dirname(__file__))` to put the log file in the project directory
2. **[Solved]** "chromedriver.exe needs to be in PATH [...]"
    * the code doesn't work when chromedriver location is specified using absolute path (no matter where chromedriver.exe is placed) => changing the code to relative path works
```python=1
# old code (doesn't work for whatever reason)
driver = webdriver.Chrome(os.path.split(os.path.realpath(__file__))[0] + r'\chromedriver.exe', options = chrome_options)
# new code (magically works!)
driver = webdriver.Chrome('chromedriver.exe', options = chrome_options)
```
3. **[Solved]** Yahoo Mall (超級商城) 只有查詢7/50個品項就結束
    * idk how I fixed it, maybe by resolving some of those exceptions?
```
# shopCrawlers.log
比對：
杜蕾斯超薄裝更薄型保險套10入 ___
現貨！Durex 杜蕾斯 衛生套 保險套 超薄裝 更薄型10入盒裝 台灣公司貨#捕夢網
INFO:root:其中一方品名.split() 的第一段子字串不是中文，拿掉重新比較
INFO:root:prod=杜蕾斯超薄裝更薄型保險套10入 ___
INFO:root:found=durex 杜蕾斯 衛生套 保險套 超薄裝 更薄型10入盒裝 台灣公 捕夢網
INFO:root:相似度大於 0.5 ，取走品名開頭非中文字串，再重新比較一次
ERROR:root:Error in Ymall!
Traceback (most recent call last):
  File "d:\eloy wu\documents\suntrail\自動查最低價\modules\crawlers\Ymall.py", line 67, in crawler_on_Ymall
    if is_same_prod(prod, found_names[i],color,threeC):
  File "d:\eloy wu\documents\suntrail\自動查最低價\modules\utilities.py", line 98, in is_same_prod
    return is_same_prod(prod[prod.index(' ') + 1:], found, color)
TypeError: is_same_prod() missing 1 required positional argument: 'threeC'
INFO:root:Yahoo!超級商城 抓取完畢！```