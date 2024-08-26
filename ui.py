from tkinter import *
from tkinter.ttk import *
import keyboard as kb
import threading


def thread_it(func, *args):
    t = threading.Thread(target=func, args=args)
    t.daemon = True
    t.start()


class WinGUI(Tk):
    def __init__(self):
        super().__init__()
        self.__win()
        self.tk_frame_managePanel = self.__tk_frame_managePanel(self)
        self.tk_button_beginButton = self.__tk_button_beginButton(self.tk_frame_managePanel)
        self.tk_frame_sheet = self.__tk_frame_sheet(self)
        self.tk_table_sheet = self.__tk_table_sheet(self.tk_frame_sheet)

    def __win(self):
        self.title("Tkinter布局助手")
        # 设置窗口大小、居中
        width = 1000
        height = 600
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        geometry = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.geometry(geometry)

        self.resizable(width=False, height=False)

    def scrollbar_autohide(self, vbar, hbar, widget):
        """自动隐藏滚动条"""

        def show():
            if vbar: vbar.lift(widget)
            if hbar: hbar.lift(widget)

        def hide():
            if vbar: vbar.lower(widget)
            if hbar: hbar.lower(widget)

        hide()
        widget.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Enter>", lambda e: show())
        if vbar: vbar.bind("<Leave>", lambda e: hide())
        if hbar: hbar.bind("<Enter>", lambda e: show())
        if hbar: hbar.bind("<Leave>", lambda e: hide())
        widget.bind("<Leave>", lambda e: hide())

    def v_scrollbar(self, vbar, widget, x, y, w, h, pw, ph):
        widget.configure(yscrollcommand=vbar.set)
        vbar.config(command=widget.yview)
        vbar.place(relx=(w + x) / pw, rely=y / ph, relheight=h / ph, anchor='ne')

    def h_scrollbar(self, hbar, widget, x, y, w, h, pw, ph):
        widget.configure(xscrollcommand=hbar.set)
        hbar.config(command=widget.xview)
        hbar.place(relx=x / pw, rely=(y + h) / ph, relwidth=w / pw, anchor='sw')

    def create_bar(self, master, widget, is_vbar, is_hbar, x, y, w, h, pw, ph):
        vbar, hbar = None, None
        if is_vbar:
            vbar = Scrollbar(master)
            self.v_scrollbar(vbar, widget, x, y, w, h, pw, ph)
        if is_hbar:
            hbar = Scrollbar(master, orient="horizontal")
            self.h_scrollbar(hbar, widget, x, y, w, h, pw, ph)
        self.scrollbar_autohide(vbar, hbar, widget)

    def __tk_frame_managePanel(self, parent):
        frame = Frame(parent, )
        frame.place(x=0, y=0, width=1000, height=140)
        return frame

    def __tk_button_beginButton(self, parent):
        btn = Button(parent, text="开始下载", takefocus=False, )
        btn.place(x=20, y=20, width=960, height=100)
        return btn

    def __tk_frame_sheet(self, parent):
        frame = Frame(parent, )
        frame.place(x=0, y=140, width=1000, height=460)
        return frame

    def __tk_table_sheet(self, parent):
        # 表头字段 表头宽度
        columns = {"标题": 575, "BV号（av号）/分p号": 383}
        tk_table = Treeview(parent, show="headings", columns=list(columns), )
        for text, width in columns.items():  # 批量设置列属性
            tk_table.heading(text, text=text, anchor='center')
            tk_table.column(text, anchor='center', width=width, stretch=False)  # stretch 不自动拉伸

        tk_table.place(x=20, y=20, width=960, height=420)
        self.create_bar(parent, tk_table, True, True, 20, 20, 960, 420, 1000, 460)
        return tk_table


class Win(WinGUI):
    def __init__(self, controller):
        self.ctl = controller
        super().__init__()
        self.__event_bind()
        # self.__style_config()
        self.ctl.init(self)

    def __event_bind(self):
        self.tk_button_beginButton.bind('<Button-1>', self.ctl.startDownloading)
        self.tk_table_sheet.bind('<Delete>', self.ctl.multiSelectDelete)
        kb.add_hotkey('Ctrl + C', lambda: threading.Thread(target=self.ctl.manageCopy, daemon=True).start())
        # kb.add_hotkey('Ctrl + C', self.ctl.manageCopy)

    # def __style_config(self):
    #     pass


if __name__ == "__main__":
    win = WinGUI()
    win.mainloop()
