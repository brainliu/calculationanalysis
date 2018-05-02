#-*-coding:utf8-*-
#user:brian
#created_at:2018/4/30 0:33
# file: GUIofznq.py
#location: china chengdu 610000
import Tkinter as tk
windows=tk.Tk()
windows.title(u"周凝倩的计算程序")
windows.geometry("200x100")
l = tk.Label(windows, text='OMG!this is TK!', bg='green', font=('Arial', 12), width=15,
             height=2)
l.pack()  # 安置
# l.place()

windows.mainloop()  # 循环