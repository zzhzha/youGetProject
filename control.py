from ui import Win
import re
import win32api
import win32con
import win32gui
import threading
import time
import win32clipboard
import subprocess
from tkinter import *
from tkinter import messagebox
import keyboard as kb
import requests
from bs4 import BeautifulSoup
import os
from pathvalidate import *
import configparser

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


# user32 = ctypes.windll.user32
# user = ctypes.windll.LoadLibrary('C:\\Windows\\System32\\user32.dll')

def getHeaders(refererUrl):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/127.0.0.0 Safari/537.36',
        'referer': refererUrl
    }
    return headers


def getMutipleVideoH1TitleAndSubtitle(text: str):
    multipartPattern = re.compile(r'title:\s+(.+?)\(P\d+?\.(.+?)\)')
    m = multipartPattern.findall(text)
    if m:
        return m[0][0], m[0][1]


def getSingleVideoTitle(text: str):
    multipartPattern = re.compile(r'title:\s+(.+)')
    m = multipartPattern.findall(text)
    if m:
        return m[0].strip()


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


def sanitizeFilename(filename: str) -> str:
    needFixstr = '@~%^&*()-+|:<>=.$`[]?!\''
    first = sanitize_filename(filename)
    for i in needFixstr:
        if i in first:
            first = first.replace(i, '_')
    return first


def thread_it(func, *args):
    t = threading.Thread(target=func, args=args)
    t.daemon = True
    t.start()


class Controller:
    # 导入UI类后，替换以下的 object 类型，将获得 IDE 属性提示功能
    ui: Win

    def __init__(self):
        self.rootPath = os.path.dirname(os.path.abspath(__file__))
        self.videoPath: str = self.rootPath + '\\Videos'  #保存所有视频的文件夹
        self.articlePath = self.rootPath + '\\Articles'
        if not os.path.exists(self.videoPath):
            os.makedirs(self.videoPath)
        if not os.path.exists(self.articlePath):
            os.makedirs(self.articlePath)

    def init(self, ui):
        """
        得到UI实例，对组件进行初始化配置
        """
        self.ui = ui
        self.s = requests.Session()
        self.cookiesPath = self.getCookiesPath()
        self.ui.protocol("WM_DELETE_WINDOW", self.closeAction)

    def addConfigUIFunction_WarnningAlert(self):
        pass

    # 获取cookies路径
    def getCookiesPath(self):
        cf = configparser.ConfigParser()
        cf.read('config.ini', encoding='utf-8')
        cookiesPath = cf.get('Path', 'cookiesPath')
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
                print(self.ui.tk_table_sheet.item(i)['values'])
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

    def initClipboard(self):
        win32clipboard.OpenClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, 'init')
        win32clipboard.CloseClipboard()
        """
        初始化剪贴板监听
        """

    # 设置剪贴板多选删除操作
    def multiSelectDelete(self, event):
        for item in self.ui.tk_table_sheet.selection():
            self.ui.tk_table_sheet.delete(item)
        """
        多选删除
        """

    # 获取剪贴板内容
    def get_clipboard_text(self):
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        return data

    def monitorClipboard(self):
        """
        监听剪贴板，当内容改变时，调用manageCopy方法
        """
        previous_content = self.get_clipboard_text()
        print(f"初始剪贴板内容: {previous_content}")
        while True:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                current_content = self.get_clipboard_text()
                if current_content != previous_content:
                    previous_content = current_content
                    thread_it(self.manageCopy)
            time.sleep(0.05)

    # 设置警告窗口置顶
    def setWarnningWindowTop(self):
        a = 1
        print("输入已阻止")
        kb.remove_all_hotkeys()

        while not win32gui.FindWindow(None, "错误"):
            time.sleep(0.05)
        else:
            print("找到错误窗口")
            hwnd = win32gui.FindWindow(None, "错误")
        while win32gui.IsWindow(hwnd):
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            if not win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

            # 以下提示SetForegroundWindow在非主线程的其他线程上调用时可能会报错（实际就是会报错）
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print('错误次数', f'{a}')
                a += 1

            # 以下提示win32con没有SC_TOPMOST属性
            # win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
            # win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_MOVE, 0)
            # win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_TOPMOST, 0)

            # win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

            time.sleep(0.05)

        print("窗口已关闭")
        kb.add_hotkey('Ctrl + C', self.manageCopy)

        # while True:
        #     try:
        #         hwnd = win32gui.FindWindow(None, "错误")
        #         win32gui.SetForegroundWindow(hwnd)
        #         time.sleep(0.5)
        #         # hwnd = win32gui.FindWindow('python.exe', None)
        #         # 窗口需要正常大小且在后台，不能最小化
        #         # win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
        #         # win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
        #         #                       win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE)
        #         # #
        #     except Exception as e:
        #         break

    # 开始下载
    def startDownloading(self, evt):
        kb.remove_all_hotkeys()
        self.ui.tk_button_beginButton.configure(state=DISABLED)
        # treeveiw没有state属性
        # self.ui.tk_table_sheet.configure(state=DISABLED)
        self.ui.tk_table_sheet.configure(selectmode='none')

        allLinksList = self.getAllDownloadLinks()
        if not allLinksList:
            thread_it(win32api.MessageBox, 0, "表中无数据，请先添加数据", '错误', win32con.MB_ICONWARNING)
            thread_it(self.setWarnningWindowTop)
            self.ui.tk_button_beginButton.configure(state=NORMAL)
            self.ui.tk_table_sheet.configure(selectmode='extended')
            return
            # treeveiw没有state属性
        '''
        [
            [('title','BV1234567890')],
            [('title','BV1234567890'),[('subtitle1','p1'),('subtitle2','p2'),('subtitle3','p3'),……]],
            [('title','cv1234567890'),['url1.jpg','url2.jpg','url3.jpg',……]]
        ]'''
        self.ui.tk_table_sheet.delete(*self.ui.tk_table_sheet.get_children())

        """
        开始下载
        """
        print("开始下载\n")
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
                    command = f'you-get -o "{saveVideoFolderPath}" -O "{generateTitle}.mp4" {videoUrl} -c {self.cookiesPath}'
                    proc = subprocess.Popen(
                        command,
                        stdin=None,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        shell=True)
                    info = proc.communicate()
                    outinfo, errinfo = info[0].decode('utf-8'), info[1].decode('utf-8')  # 获取错误信息
                    print(f"标题为：{generateTitle} 的视频下载完成")
                else:
                    # 多p视频
                    subtitlePnumberList = i[1]
                    for n in subtitlePnumberList:
                        subtitle, Pnumber = sanitizeFilename(n[0]), n[1]  #对副标题进行处理，去除特殊字符
                        if os.path.exists(os.path.join(saveVideoFolderPath, f'{subtitle}.mp4')):
                            print(f"{subtitle}.mp4已存在，跳过")
                            continue

                        videoUrl = f'https://www.bilibili.com/video/{BVnumber}?p={Pnumber[1:]}'

                        print(f"开始下载{generateTitle}下的副标题为{subtitle}的视频")
                        # command = f'you-get {videoUrl} -c {cookiesPath}'

                        command = f'you-get -o "{saveVideoFolderPath}" -O "{subtitle}.mp4" {videoUrl} -c {self.cookiesPath}'
                        print(command)
                        proc = subprocess.Popen(
                            command,
                            stdin=None,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            shell=True)
                        info = proc.communicate()
                        outinfo, errinfo = info[0].decode('utf-8'), info[1].decode('utf-8')  # 获取错误信息
                        if os.path.exists(os.path.join(saveVideoFolderPath, f'{subtitle}.mp4')):
                            print(f"{generateTitle}下的副标题为{subtitle}的视频下载完成")
                        else:
                            print(f"{generateTitle}下的副标题为{subtitle}的视频下载失败")
            else:
                saveImagesFolderPath = os.path.join(self.articlePath, generateTitle)
                if not os.path.exists(saveImagesFolderPath):
                    os.makedirs(saveImagesFolderPath)
                print(f"开始下载{generateTitle}专栏下的图片")
                for imageUrl in i[1]:
                    command = f'you-get {imageUrl} --output-dir "{saveImagesFolderPath}" -c {self.cookiesPath}'
                    proc = subprocess.Popen(
                        command,  # cmd特定的查询空间的命令
                        stdin=None,  # 标准输入 键盘
                        stdout=subprocess.PIPE,  # -1 标准输出（演示器、终端) 保存到管道中以便进行操作
                        stderr=subprocess.PIPE,  # 标准错误，保存到管道
                        shell=True)
                    info = proc.communicate()
                    outinfo, errinfo = info[0].decode('utf-8'), info[1].decode('utf-8')  # 获取错误信息
                print(f'{generateTitle}专栏下的图片全部下载完成')

        print("下载完成")

        self.ui.tk_button_beginButton.configure(state=NORMAL)
        # treeveiw没有state属性
        # self.ui.tk_frame_sheet.configure(state=NORMAL)
        self.ui.tk_table_sheet.configure(selectmode='extended')
        kb.add_hotkey('Ctrl + C', self.manageCopy)
        # 开始下载   下载完成后，调用win32api.MessageBox方法弹出提示框
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
            if imgUrlList:
                # 若解析成功，将解析结果添加到表格中
                self.addArticleItem(title, linkSign, imgUrlList)
            else:
                # 若解析失败，弹出警告窗口
                thread_it(win32api.MessageBox, 0, "链接解析失败，请检查链接是否正确", '错误', win32con.MB_ICONWARNING)
                thread_it(self.setWarnningWindowTop)

        else:
            thread_it(win32api.MessageBox, 0, "复制格式错误，请重新复制", '错误', win32con.MB_ICONWARNING)
            thread_it(self.setWarnningWindowTop)

    # 获取视频错误信息和输出信息
    def getVideoInfo(self, url, multipart=False):

        if not multipart:
            command = f'you-get -i {url} -c {self.cookiesPath}'
        else:
            command = f'you-get -i --playlist {url} -c {self.cookiesPath}'

        proc = subprocess.Popen(
            command,  # cmd特定的查询空间的命令
            stdin=None,  # 标准输入 键盘
            stdout=subprocess.PIPE,  # -1 标准输出（演示器、终端) 保存到管道中以便进行操作
            stderr=subprocess.PIPE,  # 标准错误，保存到管道
            shell=True)
        info = proc.communicate()
        outinfo, errinfo = info[0].decode('utf-8'), info[1].decode('utf-8')  # 获取错误信息
        return outinfo, errinfo

    # 检查
    def videoConfirm(self, linkSign):
        # 解析视频链接，获取视频标题和副标题。
        # 若为多p视频，则获取所有p的标题和副标题，返回字典。键为h1标题，值为列表，元素为副标题。
        # 若为单p视频，则获取视频的标题，返回字典。键为h1标题，值为空列表。
        # 若解析失败，返回False。
        videoUrl = f'https://www.bilibili.com/video/{linkSign}'
        outinfo, errinfo = self.getVideoInfo(videoUrl)
        subtitleTitleDict = {}

        if 'you-get: [error] oops, something went wrong.' in errinfo:
            # 如果链接解析失败，说明该链接下没有视频，返回False
            return False
        elif 'you-get: This is a multipart video. (use --playlist to download all parts.)' in errinfo:
            # 如果为多p视频
            multipartOutinfo = self.getVideoInfo(videoUrl, multipart=True)[0]
            for i in multipartOutinfo.split('\n'):

                # title:               進撃の巨人 CG短片 by:Masashi Imagawa (P1. Attack on titan 進撃の巨人)'
                # 当视频为多p时，会有多个标题。而此时通过you-get获取到的包含标题行中，
                # title后、括号前为视频标题，括号内为集数和分p视频的标题。
                if 'title:' in i:
                    print('出现包含标题行')
                    h1Title, Subtitle = getMutipleVideoH1TitleAndSubtitle(i.strip())

                    if not subtitleTitleDict:
                        subtitleTitleDict[h1Title] = []
                    subtitleTitleDict[h1Title].append(Subtitle)
                    #     # 去除开头的标题行
                    #     # 如 進撃の巨人 CG短片 by:Masashi Imagawa (P1. Attack on titan 進撃の巨人)
                    #     removeTitleText = i.split('title:')[1].strip()
                    #
                    #     # 获取视频标题
                    #     # 如 進撃の巨人 CG短片 by:Masashi Imagawa
                    #     h1Title=removeTitleText.split('(P1.')[0].strip()
                    #
                    #     # 若为第一集，则将标题作为字典的键，并初始化空列表作为值
                    #     subtitleTitleDict[h1Title] = []

            print('获取到多p视频标题和副标题')
            return subtitleTitleDict

        else:
            subtitleTitleDict = {}
            title = getSingleVideoTitle(outinfo.strip())
            subtitleTitleDict[title] = []
            # 给出的链接为单个有效视频链接，将链接添加到表格中
            return subtitleTitleDict

    def articleConfirm(self, linkSign):
        imgUrlPattern = re.compile('"url":"(.+?)"')
        pLabelPattern = re.compile('<p.+/p>')
        imagesUrlList = []

        try:
            int(linkSign[2:])
        except ValueError:
            return False, False
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
                elements = [i.replace(r'\u002F', '/') for i in re.findall(imgUrlPattern, response.text)]
            for u in elements:

                if not re.findall(pLabelPattern, str(u)):
                    if type(u) == str:

                        imagesUrlList.append(u)
                    else:
                        imagesUrlList.append('https:' + u['data-src'])
            print(imagesUrlList)

            return title, imagesUrlList

    def addVideoItem(self, result, linkSign):
        for title, subtitleList in result.items():
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

    # win32clipboard.OpenClipboard()
    # win32clipboard.EmptyClipboard()
    # # 设置剪贴板默认值为 constant，否则若在获取剪贴版时若剪贴板为空，会报错：
    # # TypeError: Specified clipboard format is not available
    # win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, 'constant')
    #
    # while self.ui.monitorState:
    #     time.sleep(0.05)
    #     text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
    #     if text != 'constant':
    #         win32clipboard.EmptyClipboard()
    #         win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, 'constant')
    #         exactResult = exactTypeAndId(text)
    #
    #         if exactResult[0] == 'video':
    #             pass
    #         elif exactResult[0] == 'article':
    #             pass
    #         else:
    #             thread_it(win32api.MessageBox, 0, "复制格式错误，请重新复制", '错误', win32con.MB_ICONWARNING)
    #             thread_it(setWarnningWindowTop)

    # win32clipboard.CloseClipboard()

    def closeAction(self):
        if messagebox.askokcancel("Quit", "你确认要退出?"):
            self.ui.destroy()

        # self.ui.monitorState = False
        # win32clipboard.CloseClipboard()

        self.ui.quit()
