说明：

该项目基于you-get项目制作，因此实质上只是为you-get做了一个图形化的界面，和添加了一个批量选中视频和我自己添加的爬取bilibili专栏图片的功能，范围还只限定在bilibili。
作为刚入门编程的学生，由于能力有限，程序设计得非常粗糙，例如不清楚怎么实现模态窗口，最后用了很奇怪的方法实现相似的效果，并且可能还含有bug。


使用前需要干的事：

使用前需要下载火狐浏览器，并在b站上登录账号。接着获取火狐浏览器软件生成的cookies.sqlite（找不到可以使用Everything软件直接搜索），复制它的完整路径到config.ini里（或者直接把cookies.sqlite文件本身复制到软件文件夹下，直接填cookies.sqlite）。


使用方式：

运行程序后，程序会自动检测Ctrl+c复制快捷键的触发，提取剪贴板内的视频的BV号或专栏的cv号（注意通过使用右键复制等操作不会被程序捕获）
复制时不用精确到BV号（cv号）。复制整个网址url都行，只要其中包含格式正确且有效的BV号（cv号）就行。
复制错误有弹窗提示，并阻止下一步复制操作直到关闭弹窗。
多p视频和专栏可以双击视频或专栏标题项折叠；按住Ctrl单击左键可以选中多项，按delete键删除不需要的视频或图片。
下载的视频在Videos文件夹里，专栏图片在Articles文件夹里。（这两个文件夹在第一次使用时会自动生成）


注意：

下载的视频都是账号所允许的最高画质（包括图片），且下载一旦开始便不能暂停，因此需注意是否预留有足够的空间大小。

最后我还在发行版乱画了一个图标把它加进去（）



