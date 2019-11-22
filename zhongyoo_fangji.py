# coding=gbk
import requests
import lxml.html
import os
import time

ori_url = 'http://www.zhongyoo.com/fangji/page_'
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'Connection': 'keep-alive',
    'Cookie': '__cfduid=db3aacee147f322c6d9dde3d7e9c7b11f1570152723; Hm_lvt_f9eb7a07918590a54f0fa333419bae7e=1570152721,1570154532; Hm_lpvt_f9eb7a07918590a54f0fa333419bae7e=1570161610',
    'Host': 'www.zhongyoo.com',
    'If-Modified-Since': 'Sat, 09 Apr 2016 06:52:06 GMT',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
}

def page():
    page = []
    for i in range(1,21):
        x = ori_url+str(i)+'.html'
        page.append(x)
    return page

def herb(page_URL):  # page_URL为列表
    herb_href = []
    for x in page_URL:
        ori_info = requests.get(x).content.decode('gbk',errors='ignore')
        selector = lxml.html.fromstring(ori_info)
        info1 = selector.xpath('/html/body/div[@class="wrap"]/div[@class="con_left"]')
        info2 = info1[0].xpath('//div[@class="listbox"]/ul/li/a')
        for i in info2:
            href = i.xpath('@href')  #/html/body/div[5]/div[2]/div[2]/ul[1]/li[1]/a
            # href = href[16:]
            for m in href:
                herb_href.append(m)
        # herb_href = herb_href[16:]
    return herb_href

def get_herbinfo(herb_url):
    info_text = []
    # herb_url= ','.join(herb_url)
    herb_ori = requests.get(herb_url,headers=headers).content.decode('gbk',errors='ignore')
    selector = lxml.html.fromstring(herb_ori)
    name = selector.xpath('//div[@class="wrap"]/div[@class="con_left"]/div[@class="art_box"]/h1/text()')
    herb_text0 = selector.xpath('//div[@class="wrap"]/div[@class="con_left"]/div[@class="art_box"]/div[@class="art_con"]/p')
    for i in herb_text0:
        herb_text = i.xpath('string(.)')
        info_text.append(herb_text)
    return name,info_text

def save(herbname,herbinfo):
    herbinfo = ','.join(herbinfo)
    os.makedirs('zhongyoo_TCM',exist_ok=True)
    with open(os.path.join(herbname+'.txt'),'w',encoding='gbk',errors='ignore') as f:
        f.write(herbinfo)


if __name__ == '__main__':
    page_url = page()
    herb_url = herb(page_url)
    for i in herb_url:
        name, info = get_herbinfo(i)
        save(name[0],info)
        time.sleep(0.5)
