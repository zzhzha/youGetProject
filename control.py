import json
from ui import Win
import re
import win32api
import win32con
import win32gui
import threading
import time
import win32clipboard
from tkinter import *
from tkinter import messagebox
import keyboard as kb
import sys
import requests
from bs4 import BeautifulSoup
import os
from pathvalidate import *
import configparser
from subprocess import run


"""

数据存储格式：
所有数据存贮在列表中，该列表内的每个列表元素代表一个视频或文章。

视频或文章总共包含以下几个要素：
视频：
1. 标题
2. 副标题（如果为多p视频）
3. BV号
下载和解析视频时，需要用到BV号和p号；保存时需要用到标题和副标题。
单个视频数据采用以下存贮方式：
[('title','BV1234567890')]

多p视频数据采用以下存贮方式：
[('title','BV1234567890'),[('subtitle1','p1'),('subtitle2','p2'),('subtitle3','p3'),……]]


文章：
1. 标题
2. 文章号（cv号）
3. 图片链接
[('title','cv1234567890'),['url1.jpg','url2.jpg','url3.jpg',……]]


[
[('title','BV1234567890')],
[('title','BV1234567890'),[('subtitle1','p1'),('subtitle2','p2'),('subtitle3','p3'),……]],
[('title','cv1234567890'),['url1.jpg','url2.jpg','url3.jpg',……]]
]

"""


def getHeaders(refererUrl='www.bilibili.com'):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/127.0.0.0 Safari/537.36',
        'referer': refererUrl
    }
    return headers


def exactTypeAndLinkSign(url: str) -> [str, str]:
    # 初步判断链接类型是否符合视频或文章
    # 视频BV号：dfghdvbnm/BV1234567890/?fsdfsdf
    # 多p视频BV号：
    # sdfgsdfgxcvvcbn/BV17V4y1n7v5?p=1sdfsdfsdfsdf
    # vnhfgjytuiyu/BV17V4y1n7v5/?p=2getysrtgsdfg

    # 文章cv号：cv1234567890
    videoSymbolPattern = re.compile('BV.+')
    articleSymbolPattern = re.compile(r'cv\d+')
    for i in url.split('/'):
        if videoSymbolPattern.search(i):
            # BV17V4y1n7v5?p=1sdfsdfsdfsdf
            for n in i.split('?'):
                if videoSymbolPattern.search(n):
                    return ['video', n]
        if articleSymbolPattern.search(i):
            return ['article', i]
    return ['', 'None']


def thread_it(func, *args, daemon: bool = True):
    t = threading.Thread(target=func, args=args)
    t.daemon = daemon
    t.start()


class Controller:
    # 导入UI类后，替换以下的 object 类型，将获得 IDE 属性提示功能
    ui: Win

    def __init__(self):
        self.rootPath = os.getcwd()
        self.videoPath: str = self.rootPath + '\\Videos'
        self.articlePath = self.rootPath + '\\Articles'
        self.cookiesConfigIniFile = self.rootPath + '\\config.ini'
        self.VideoInfoConfirmPattern = re.compile("<script>window\.__INITIAL_STATE__=(\{.*?});\(function")

        if not os.path.exists(self.videoPath):
            os.makedirs(self.videoPath)
        if not os.path.exists(self.articlePath):
            os.makedirs(self.articlePath)
        if not os.path.exists(self.cookiesConfigIniFile):
            with open(self.cookiesConfigIniFile, 'w', encoding='utf-8') as f:
                f.write('[Path]\ncookiesPath=')

    def init(self, ui):
        """
        得到UI实例，对组件进行初始化配置
        """
        self.ui = ui
        self.s = requests.Session()
        self.cookiesPath = self.getCookiesPath()
        self.ui.protocol("WM_DELETE_WINDOW", self.closeAction)

    # 获取cookies路径
    def getCookiesPath(self):
        cf = configparser.ConfigParser()
        cf.read(self.cookiesConfigIniFile, encoding='utf-8')
        cookiesPath = cf.get('Path', 'cookiesPath')
        if not os.path.exists(cookiesPath):
            thread_it(win32api.MessageBox, 0, "请先在ini填写cookies路径", '错误', win32con.MB_ICONWARNING, daemon=False)
            sys.exit()
        return cookiesPath

    # 获取需要下载的所有的链接
    def getAllDownloadLinks(self):
        """
        获取所有下载链接
        """
        links = []
        #
        for i in self.ui.tk_table_sheet.get_children():
            # 获取所有子节点
            title, link = self.ui.tk_table_sheet.item(i)['values'][0], self.ui.tk_table_sheet.item(i)['values'][1]

            if not self.ui.tk_table_sheet.get_children(i):
                # 如果子节点没有子节点，说明为单个视频，则获取标题和BV号
                # [('title','BV1234567890')]
                # print(self.ui.tk_table_sheet.item(i)['values'])
                videoItem = [(title, link)]
                links.append(videoItem)

            else:
                # 如果子节点有子节点，说明为多p视频或文章，则获取标题、BV号，和子节点标题和P号或cv号和图片链接
                if 'BV' in link:
                    # 多p视频
                    # [('title','BV1234567890'), [('subtitle1','p1'), ('subtitle2','p2'), ('subtitle3','p3'),……]]

                    multipleVideoItem = []
                    multipleVideoItem.append((title, link))
                    multipleVideoItem.append([])
                    for n in self.ui.tk_table_sheet.get_children(i):
                        subtitle, Pnumber = self.ui.tk_table_sheet.item(n)['values'][0], \
                            self.ui.tk_table_sheet.item(n)['values'][1]
                        multipleVideoItem[1].append((subtitle, Pnumber))

                    links.append(multipleVideoItem)
                else:
                    # 文章
                    # [('title','cv1234567890'),['url1.jpg','url2.jpg','url3.jpg',……]]
                    articleItem = []
                    articleItem.append((title, link))
                    articleItem.append([])
                    for n in self.ui.tk_table_sheet.get_children(i):
                        imgUrl = self.ui.tk_table_sheet.item(n)['values'][1]
                        articleItem[1].append(imgUrl)

                    links.append(articleItem)

        return links

    # 设置剪贴板多选删除操作
    def multiSelectDelete(self, event):
        """
        多选删除
        """
        selected_items = self.ui.tk_table_sheet.selection()  # 获取所有选中的项的ID
        for i in self.ui.tk_table_sheet.get_children():
            for j in selected_items:
                if j == i:
                    self.ui.tk_table_sheet.delete(i)
                else:
                    if j in self.ui.tk_table_sheet.get_children(i):
                        self.ui.tk_table_sheet.delete(j)

    def setWarnningWindowTop(self):
        a = 1
        print("输入已阻止")
        self.ui.tk_button_beginButton.configure(state=DISABLED)
        self.ui.tk_table_sheet.configure(selectmode='none')
        kb.remove_all_hotkeys()

        while not win32gui.FindWindow(None, "错误"):
            time.sleep(0.05)
        else:
            print("找到错误窗口")
            hwnd = win32gui.FindWindow(None, "错误")
        # 目标是弹出窗口后用户不能做出任何操作直到关闭窗口。并且该窗口一直保持置顶。
        # 因此窗口需要一直处于焦点状态
        while win32gui.IsWindow(hwnd):
            if not win32gui.IsWindowVisible(hwnd):
                # 若窗口不可见（如表现为窗口被最小化到任务栏，窗口被程序主动隐藏，窗口被其他窗口覆盖（不在最上层
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print('错误次数', f'{a}')
                a += 1
            time.sleep(0.05)

        print("窗口已关闭")
        self.ui.tk_table_sheet.configure(selectmode='extended')
        self.ui.tk_button_beginButton.configure(state=NORMAL)
        kb.add_hotkey('Ctrl + C', lambda: threading.Thread(target=self.manageCopy, daemon=True).start())

    # 开始下载
    def startDownloading(self):
        kb.remove_all_hotkeys()
        self.ui.tk_button_beginButton.configure(state=DISABLED)
        self.ui.tk_table_sheet.configure(selectmode='none')
        allLinksList = self.getAllDownloadLinks()
        self.ui.tk_table_sheet.delete(*self.ui.tk_table_sheet.get_children())

        if not allLinksList:
            thread_it(win32api.MessageBox, 0, "表中无数据，请先添加数据", '错误', win32con.MB_ICONWARNING)
            # 上面去除的快捷键等在setWarnningWindowTop中重新设置
            thread_it(self.setWarnningWindowTop)
            return
        '''
        [
            [('title','BV1234567890')],
            [('title','BV1234567890'),[('subtitle1','p1'),('subtitle2','p2'),('subtitle3','p3'),……]],
            [('title','cv1234567890'),['url1.jpg','url2.jpg','url3.jpg',……]]
        ]'''

        """
        开始下载
        """
        for i in allLinksList:
            generateTitle = sanitize_filename(i[0][0])  #对标题进行处理，去除特殊字符
            if 'BV' in i[0][1]:
                BVnumber = i[0][1]
                saveVideoFolderPath = os.path.join(self.videoPath, generateTitle)  #视频保存的文件夹

                if not os.path.exists(saveVideoFolderPath):
                    os.makedirs(saveVideoFolderPath)

                if len(i) == 1:
                    # 单个视频
                    videoPath = os.path.join(saveVideoFolderPath, f'{generateTitle}.mp4')
                    print(videoPath)
                    # 以视频标题为文件夹名，创建视频保存的文件夹
                    if os.path.exists(videoPath):
                        # 视频已存在，跳过
                        print(f"{generateTitle}已存在，跳过")
                        continue

                    videoUrl = f'https://www.bilibili.com/video/{BVnumber}'

                    print(f"开始下载：{generateTitle}")
                    command = f'you-get --skip-existing-file-size-check -o "{saveVideoFolderPath}" -O "{generateTitle}" {videoUrl} -c {self.cookiesPath}'
                    thread_it(lambda arg:run(arg, shell=True), command, daemon=False)

                else:
                    # 多p视频
                    subtitlePnumberList = i[1]
                    for n in subtitlePnumberList:
                        subtitle, Pnumber = sanitize_filename(n[0]), n[1]  #对副标题进行处理，去除特殊字符
                        #多p视频文件格式：总标题(PX.分p标题).mp4
                        videoPath = os.path.join(saveVideoFolderPath, f'{subtitle}.mp4')
                        if os.path.exists(videoPath):
                            # 视频已存在，跳过
                            print(f"{subtitle}.mp4已存在，跳过")
                            continue
                        videoUrl = f'https://www.bilibili.com/video/{BVnumber}?p={Pnumber[1:]}'

                        print(f"开始下载{generateTitle}下的副标题为{subtitle}的视频")
                        command = f'you-get --skip-existing-file-size-check -o "{saveVideoFolderPath}" -O "{subtitle}" {videoUrl} -c {self.cookiesPath}'
                        thread_it(lambda arg:run(arg, shell=True), command, daemon=False)

            else:
                saveImagesFolderPath = os.path.join(self.articlePath, generateTitle)
                if not os.path.exists(saveImagesFolderPath):
                    os.makedirs(saveImagesFolderPath)
                print(f"开始下载{generateTitle}专栏下的图片")
                for imageUrl in i[1]:
                    command = f'you-get {imageUrl} --skip-existing-file-size-check --output-dir "{saveImagesFolderPath}" -O "{i[1].index(imageUrl) + 1}" -c {self.cookiesPath}'
                    thread_it(lambda arg:run(arg, shell=True), command, daemon=False)
                    time.sleep(0.1)

        kb.add_hotkey('Ctrl + C', lambda: threading.Thread(target=self.manageCopy, daemon=True).start())
        self.ui.tk_button_beginButton.configure(state=NORMAL)
        self.ui.tk_table_sheet.configure(selectmode='extended')
        time.sleep(0.05)

    # 获取到剪贴板内容改变后，做出相应的操作
    def manageCopy(self):
        print("剪贴板内容改变")
        # 需要进行延迟，否则会导致剪贴板获取到上一次的剪贴板内容或空。
        # 而当前一次的剪贴板内容为空时，调用win32clipboard.GetClipboardData会报错
        time.sleep(0.1)
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        type, linkSign = exactTypeAndLinkSign(text)
        # 若复制的链接格式不正确，弹出警告窗口
        if type == 'video':
            result = self.videoConfirm(linkSign)

            if result:
                # 若解析成功，将解析结果添加到表格中
                self.addVideoItem(result, linkSign)

            else:
                # 若解析失败，弹出警告窗口
                thread_it(win32api.MessageBox, 0, "链接解析失败，请检查链接是否正确", '错误', win32con.MB_ICONWARNING)
                thread_it(self.setWarnningWindowTop)
        elif type == 'article':
            title, imgUrlList = self.articleConfirm(linkSign)
            if isinstance(imgUrlList, list):
                self.addArticleItem(title, linkSign, imgUrlList)

            else:
                if imgUrlList:
                    # imgUrlList为True，说明图片链接列表中有重复项，直接返回
                    return
                else:
                    # 若解析失败，弹出警告窗口
                    thread_it(win32api.MessageBox, 0, "链接解析失败，请检查链接是否正确", '错误',
                              win32con.MB_ICONWARNING)
                    thread_it(self.setWarnningWindowTop)

        else:
            thread_it(win32api.MessageBox, 0, "复制格式错误，请重新复制", '错误', win32con.MB_ICONWARNING, )
            thread_it(self.setWarnningWindowTop)

    # 获取视频错误信息和输出信息
    def getVideoInfo(self, videoUrl):
        '''
        获取视频信息,返回response对象,若获取失败,返回False
        :param videoUrl:
        :return: response，若获取失败，返回False
        '''
        response = self.s.get(videoUrl, headers=getHeaders())
        response.encoding = 'utf-8'
        if response.status_code == 200:
            return response
        else:
            return False

    def videoConfirm(self, linkSign):
        # 解析视频链接，获取视频标题和副标题。
        # 若为多p视频，则获取所有p的标题和副标题，返回字典。键为h1标题，值为列表，元素为副标题。
        # 若为单p视频，则获取视频的标题，返回字典。键为h1标题，值为空列表。
        # 若解析失败，返回False。
        videoUrl = f'https://www.bilibili.com/video/{linkSign}'
        start_time = time.time()
        response = self.getVideoInfo(videoUrl)
        # outinfo, errinfo = self.getVideoInfo(videoUrl)
        end_time = time.time()
        print(f"解析视频链接耗时：{end_time - start_time}秒")
        '''        
        格式为 {}，键为h1标题，值为列表，元素为副标题。
        若为单p视频，则返回 { h1标题 : [] } 
        若为多p视频，则返回 { h1标题 : [ 副标题1, 副标题2, 副标题3, …  ] }
        '''

        if not response:
            # 如果链接解析失败，说明该链接下没有视频，返回False
            return False
        else:
            # 如果为多p视频
            subtitleTitleDict = {}
            text = self.VideoInfoConfirmPattern.search(response.text).group(1)
            exactInfoFix = re.sub(r'(\\u[a-zA-Z0-9]{4})', lambda x: x.group(1).encode("utf-8").decode("unicode_escape"),
                                  text)
            exactInfoJson = json.loads(exactInfoFix)
            videoData = exactInfoJson['videoData']
            title = videoData['title']
            subtitleTitleDict[title] = []
            videosInfoLIist = videoData['pages']
            if len(videosInfoLIist) == 1:
                # 单p视频
                subtitleTitleDict[title] = []
                return subtitleTitleDict

            else:
                # 多p视频
                for i in range(len(videosInfoLIist)):
                    subtitle = videosInfoLIist[i]['part']
                    subtitleTitleDict[title].append(subtitle)
                return subtitleTitleDict

    def articleConfirm(self, linkSign):
        getImgUrlPattern = re.compile('"url":"(.+?)"')
        pLabelPattern = re.compile('<p.+/p>')
        finalImgUrlConfirm = re.compile('https://i[0-9]+\.hdslb\.com/bfs/article/.*')
        imagesUrlList = []

        try:
            int(linkSign[2:])
        except ValueError:
            return False, False
        for i in self.ui.tk_table_sheet.get_children():
            if self.ui.tk_table_sheet.item(i)['values'][1] == linkSign:
                # 若专栏已存在，则跳过
                return True, True

        articleUrl = f'https://www.bilibili.com/read/{linkSign}'
        headers = getHeaders(articleUrl)
        response = self.s.get(articleUrl, headers=headers)
        response.encoding = 'utf-8'
        if '什么都没有找到' in response.text:
            return False, False
        else:
            soup = BeautifulSoup(response.text, 'lxml')
            title = soup.find('title').text.strip()
            try:
                # 获取所有img标签和p标签
                elements = soup.select_one('#read-article-holder').find_all(['img', 'p'])
            except Exception:
                elements = [i.replace(r'\u002F', '/') for i in re.findall(getImgUrlPattern, response.text)]
            for u in elements:

                if not re.findall(pLabelPattern, str(u)):
                    if type(u) == str:

                        imagesUrlList.append(u)
                    else:
                        initImagesUrl = u['data-src']
                        if re.findall(finalImgUrlConfirm, initImagesUrl):
                            imagesUrlList.append(initImagesUrl)
                        else:
                            imagesUrlList.append('https:' + initImagesUrl)

            return title, imagesUrlList

    def addVideoItem(self, result, linkSign):
        for title, subtitleList in result.items():
            for i in self.ui.tk_table_sheet.get_children():
                if self.ui.tk_table_sheet.item(i)['values'][0] == title:
                    # 若标题已存在，则跳过
                    return
            if not subtitleList:
                # 视频为单p，则将标题和BV号添加到表格中
                self.ui.tk_table_sheet.insert('', END, open=True, text='', values=(title, linkSign))
            else:
                # 视频为多p，则将标题和BV号添加到表格中，并将所有p的副标题和在视频中的顺序（如P1、P2、P3）添加到标题节点下中
                fatherNode = self.ui.tk_table_sheet.insert('', END, open=True, text='',
                                                           values=(title, linkSign))
                # 给每个标题节点下添加子节点，每个子节点包含一个副标题和在视频中的顺序

                # 如：P1、P2、P3

                for subtitle in subtitleList:
                    self.ui.tk_table_sheet.insert(fatherNode, END, open=True, text='',
                                                  values=(subtitle, f'P{subtitleList.index(subtitle) + 1}'))

    def addArticleItem(self, title, linkSign, imgUrlList):
        # 标题节点下添加子节点，每个子节点包含一个图片链接,第一个值为空，第二个值为图片链接
        fatherNode = self.ui.tk_table_sheet.insert('', END, open=False, text='', values=(title, linkSign))
        for imgUrl in imgUrlList:
            self.ui.tk_table_sheet.insert(fatherNode, END, open=True, text='', values=(' ', imgUrl))

    def closeAction(self):
        if messagebox.askokcancel("Quit", "你确认要退出?"):
            self.ui.destroy()

        self.ui.quit()
