import requests
from bs4 import BeautifulSoup
import os
import re

title = ''

def getMusic(ID,path):
    '''
    下载单个歌曲

    :param ID: 歌曲的id信息
    :param path: 歌曲的下载路径
    :return: None
    '''
    try:
        kv = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.100 Safari/537.36'}
        cloud = "http://music.163.com/song/media/outer/url?id="
        url=cloud+ID+".mp3"
        tmp = requests.get(url, headers=kv)
        tmp.raise_for_status()
        # tmp.encoding=tmp.apparent_encoding 不用解码 因为是二进制文件
        print("访问成功，正在下载，请稍后......")
        with open(path,"wb") as f:
            f.write(tmp.content)
        f.close()
        print("下载成功")
    except:
        print("访问错误")
        print("请确认你的网络连接或者输入id是否正确")


def getMusicList(ID):
    '''
    获取歌单的信息--歌曲名字--歌曲的id
    :param ID: 歌单的id
    :return: 存放歌曲名字和id的字典
    '''
    headers = {
        'Referer': 'http://music.163.com/',
        'Host': 'music.163.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.75 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    }
    #请求头 伪装成用户访问

    base_url = 'http://music.163.com/playlist?id='
    s = requests.session()#保持会话
    response = s.get(base_url+str(ID), headers=headers).content
    ss = BeautifulSoup(response, 'lxml')
    #使用bs4解析o(￣ヘ￣o＃)

    main = ss.find('ul', {'class': 'f-hide'})
    ls = main.find_all('a')
    #ls 是迭代器 存放所有歌的标签

    global title
    title = ss.find('h2', {'class': 'f-ff2 f-brk'}).string      #获取歌单的名字
    print('一共{}首歌'.format(len(ls)))
    Total_dic = {'title':title}
    Music_dic = {}
    #声明字典

    for music in ls:
        Mname = music.text
        MID = str(music['href']).replace('/song?id=', '')
        print('Name : {:<30} \tID : {:^10}'.format(Mname, MID))
        Music_dic[Mname] = MID
        #存放进字典 key是name value是id
    Total_dic['songlist'] = Music_dic
    return Total_dic
    #返回字典


def getID():
    '''
    输入程序 利用正则表达式判断用户输入是否是正确的id
    :return: 歌单的id，以及bool值
    '''
    playlist_ID = input("请输入下载歌单的id : ")
    pattern = re.compile(r'\d+')
    ls = re.findall(pattern, playlist_ID)
    if len(ls)!= 0 :
        return playlist_ID, True
    else:
        return 0, False

def main():
    '''
    主函数 执行下载
    :return: None
    '''
    while True:
        ID, flag = getID()
        if flag:
            listJson = getMusicList(ID)
            # 创建文件夹
            if os.path.exists(title):
                print('歌曲目录文件夹已存在')
            else:
                os.mkdir(title)
                print('已创建歌曲目录文件夹')
            dic = listJson['songlist']
            for name in dic.keys():
                print(name, end='  ')
                getMusic(dic[name], title + '/' + name + '.mp3')
                #下载每一首歌 第二个参数是路径 必须要加".mp3"
            break
        else:
            print('请输入歌单的id， 您有可能错误输入了歌单的网址')

if __name__=='__main__':
    main()