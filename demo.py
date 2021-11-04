import hashlib
import math
import os
import sys
import time
import tkinter

import xlrd
import xlwt
import openpyxl

#打开游戏

# def openGame():
#   url = "F:\龙武\LongWuWD\LWLaunch.exe"
#   #运行这个文件
#   os.popen(url)
#   time.sleep(3)
#   print("打开游戏成功！")
#test insert
#获取当前目录所有.exe安装包
dir = []


def getfile():
        path = "F:/file"  #F盘创建一个file文件夹，安装包放里面
        allfile = os.listdir(path)  #listdir列出path下的所有文件名，以列表的形式返回
        for f in allfile:
            if os.path.splitext(f)[1] == '.exe':  #只读.exe结尾的文件
                dir.append(f)
        print(dir)
#获取当前目录所有.exe安装包大小
packsize = []


def getSize():
    try:
        for file in dir:
            path = os.path.join('F:/file', file)  #os.join用于拼接文件目录
            filesizeBT = os.path.getsize(path)
            if filesizeBT > 1073741824:
                filesizeG = round(float(filesizeBT / 1024**3), 2)  #将字节数换算成GB,保留两位小数                               
                one = str(filesizeBT)
                two = str(filesizeG)
                size = two + "GB" + ' ' + '(' + one + "字节" + ')'
                packsize.append(size)
            else:
                filesize = filesizeBT // 1024 // 1024  #将字节数换算成MB
                one = str(filesizeBT)
                two = str(filesize)
                size = two + "MB" + ' ' + '(' + one + "字节" + ')'
                packsize.append(size)
        print(packsize)
    except:
        print('你有问题？')


#计算文件hash值（md5码）
md5 = []


def getMD5():
    try:
        md5_hash = hashlib.md5()
        for filename in dir:
            filename = os.path.join('F:/file', filename)
            with open(filename, "rb") as f:
                for byte_block in iter(lambda: f.read(4096), b""):
                    md5_hash.update(byte_block)
                md5.append(md5_hash.hexdigest())
        print(md5)
    except:
        print('你又怎么了？')

#获取当前时间用于显示
timex =[]
time = time.localtime(time.time())
for time in time:
    timex.append(time)
#将信息填入表格
def wSheet():
    path = 'F:file/安装包需求{}-{}-{}.xlsx'.format(timex[0],timex[1],timex[2])
    open(file= path,mode='a+')
    table = xlrd.open_workbook()
    wb2 = load_workbook(path)
    
if __name__ == "__main__":
    getfile()
#创建窗口
window = tkinter.Tk()  #实例化一个窗口组件
window.geometry('500x500+10+10')  #geometry属性设置窗口大小
window.title("周二安装包验收信息填写 V1.0")  #左上角title设置
#openGameButton = tkinter.Button(window, text ="openGame", command = openGame)
#openGameButton.pack()
getSizeButton = tkinter.Button(window, text="取安装包大小", command=getSize, bd=5)  #实例化一个按钮控件
#getSizeButton.pack()
getSizeButton.place(relx=0, rely=0)  #place属性用于控制子组件相对于父组件的位置
getHashButton = tkinter.Button(window, text="计算hash值", command=getMD5, bd=5)
#getHashButton.pack()
getHashButton.place(relx=0, rely=0.1)
WsheetButton = tkinter.Button(window, text="信息填入表格", command=wSheet, bd=5)
#WsheetButton.pack()
WsheetButton.place(relx=0, rely=0.2)
#killGameButton = tkinter.Button(window, text = "killGame", command = killGame)
#killGameButton.pack()
#logDate = tkinter.Text(window, text = '日志',windth = 70, height = 49) #本意是想在UI界面输出日志，但为什么运行不了？
#logDate.insert(1.0, md5)
Label2 = tkinter.Label(window, text='---------------------demo-------------------').grid(padx=120,pady=470)                                                               
window.mainloop()
