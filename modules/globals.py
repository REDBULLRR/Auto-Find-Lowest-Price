import random

# ======== Global Variables =========
# - - - 如果一直用同一個User Agent去請求網站回應的話，有可能會被鎖IP，所以這邊先列出一系列隨機的備用 UserAgent - - -
# 可新增。至 https://www.toolnb.com/tools-lang-zh-TW/createuseragent.html 獲取隨機 User Agent

UserAgents = [
    'Mozilla/5.0 (Macintosh; PPC Mac OS X 10_10_9 rv:2.0; he-IL) AppleWebKit/533.12.3 (KHTML, like Gecko) Version/5.0 Safari/533.12.3',
    'Mozilla/5.0 (iPad; CPU iPad OS 5_1_1 like Mac OS X) AppleWebKit/531.0 (KHTML, like Gecko) CriOS/22.0.834.0 Mobile/87E842 Safari/531.0',
    'Mozilla/5.0 (iPad; CPU iPad OS 5_1_1 like Mac OS X) AppleWebKit/531.0 (KHTML, like Gecko) CriOS/43.0.882.0 Mobile/37C008 Safari/531.0',
    'Mozilla/5.0 (iPod; U; CPU iPhone OS 3_3 like Mac OS X; uk-UA) AppleWebKit/531.31.5 (KHTML, like Gecko) Version/4.0.5 Mobile/8B117 Safari/6531.31.5',
    'Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10_11_2 rv:5.0; dz-BT) AppleWebKit/532.45.2 (KHTML, like Gecko) Version/4.0 Safari/532.45.2',
    'Mozilla/5.0 (iPhone; CPU iPhone OS 6_1_6 like Mac OS X) AppleWebKit/532.2 (KHTML, like Gecko) FxiOS/10.9n4610.0 Mobile/30N900 Safari/532.2',
    'Mozilla/5.0 (iPod; U; CPU iPhone OS 3_0 like Mac OS X; gu-IN) AppleWebKit/533.20.7 (KHTML, like Gecko) Version/3.0.5 Mobile/8B115 Safari/6533.20.7'
]

headers = {
    'cookie': 'ECC=GoogleBot',
    # - - 使用隨機的 User-Agent - - -
    "User-Agent": UserAgents[(random.randint(0, len(UserAgents)-1))]
}
