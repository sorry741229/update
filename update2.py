#-*- coding: utf-8 -*-
import time
import os, threading
import os, sys
from watchdog.observers import Observer
from watchdog.events import *
from watchdog.utils.dirsnapshot import DirectorySnapshot, DirectorySnapshotDiff
from watchdog.events import LoggingEventHandler
from colorama import init, Fore, Back #字體顏色
init(autoreset=True)#字體顏色
import plyer
from plyer import notification
import requests
#調整視窗大小
from ctypes import windll, byref
from ctypes.wintypes import SMALL_RECT
import ctypes
import threading
import tkinter as tk
import pystray
from PIL import Image
from pystray import MenuItem, Menu


def quit_window(icon: pystray.Icon):
    icon.stop()
    win.destroy()


def show_window():
    win.deiconify()


def on_exit():
    win.withdraw()


def delete_all():
    text.delete(1.0, "end")      # 刪除全部內容




menu = (MenuItem('顯示', show_window, default=True), Menu.SEPARATOR, MenuItem('退出', quit_window))
image = Image.open('C:/Users/Public/Pictures/light.ico')
icon = pystray.Icon("icon", image, "群旭科技發行圖面通知", menu)
win = tk.Tk()
win.title('群旭科技發行圖面通知')
win.geometry("1100x200")
win.resizable(width= 1, height=False)
text = tk.Text(win)  # 顯示文字
text.config(width = 200 , height = 50)
text.config(bg = '#dcdcdc' , fg = '#191970')
scrollbar = tk.Scrollbar(win)          # 建立滾動條
scrollbar.pack(side='right', fill='y')
btn1=tk.Button(win, text="清空畫面", command=delete_all)
btn1.pack(side=tk.BOTTOM)
text.tag_configure("left", justify='left') #輸出文字對齊用


WindowsSTDOUT = windll.kernel32.GetStdHandle(-11)
dimensions = SMALL_RECT(-10, -10, 100, 20) # (left, top, right, bottom)
# Width = (Right - Left) + 1; Height = (Bottom - Top) + 1
windll.kernel32.SetConsoleWindowInfo(WindowsSTDOUT, True, byref(dimensions))



#授權時間
def now():
    return time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
s = '2023-06-14 23:59:59'
if now() > s:
    #notification.notify(title = '歡迎使用Kolink傳真通知', message = 'by F0614 Martin Chung at Coolink CNC Dept. in DEC 2022' ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 5 )
    text.insert(tk.INSERT, '歡迎使用群旭科技發行圖面通知'+ str('\n'))
    text.insert(tk.INSERT, 'by Martin Chung at Coolink CNC Dept. in DEC 2022'+ str('\n'))
else:
    text.insert(tk.INSERT, '歡迎使用群旭科技發行圖面通知'+ str('\n'))

if os.path.isfile('C:/Users/Public/Documents/發行圖面更新紀錄.csv'): #檢查檔案在不在--公司用
    #print('找到更新紀錄檔')
    text.insert(tk.INSERT, ' '+str('\n'))
    text.insert(tk.INSERT, '找到更新紀錄檔...'+str('\n'))
    text.pack()
else:
    #print('更新記錄檔不存在')
    #print('將會自動建立發行圖面更新紀錄檔')
    text.insert(tk.INSERT, ' '+str('\n'))
    text.insert(tk.INSERT, '更新記錄檔不存在'+ str('\n'))
    text.insert(tk.INSERT, '將會自動建立發行圖面更新紀錄檔...'+ str('\n'))
    text.pack()

    with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'w', encoding = 'cp950') as f:
        a = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        f.write(str(a)+ '\n')

notification.notify(title = '發行圖面通知系統', message = '啟動偵測中' ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1.5 )#--公司用

class FileEventHandler(FileSystemEventHandler):
    def __init__(self, aim_path):
        FileSystemEventHandler.__init__(self)
        self.aim_path = aim_path
        self.timer = None
        self.snapshot = DirectorySnapshot(self.aim_path)
        self.timer = threading.Timer(0.5, self.checkSnapshot)
        self.timer.start()

    
    def on_any_event(self, event):
        # if self.timer:
        #     self.timer.cancel()
        self.timer = threading.Timer(0.5, self.checkSnapshot)
        self.timer.start()
    
    def checkSnapshot(self):   
        snapshot = DirectorySnapshot(self.aim_path)
        diff = DirectorySnapshotDiff(self.snapshot, snapshot)
        self.snapshot = snapshot
        self.timer = threading.Timer(0.5, self.snapshot)
        #print(diff)
        if diff.files_created == []:
            pass
        else:
            dc_ans = []
            for dc in diff.files_created:
                log = []
                ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))," 圖面新增:  ", dc
                #print(ans[0], Fore.YELLOW + (ans[1]), Fore.CYAN + (ans[2][16:]))
                text.config(state = 'normal')
                text.insert(tk.INSERT, str((ans[0][5:]) + ans[1] + ans[2][15:] + '\n'))
                text.tag_add('left', 1.0, "end")  #輸出文字對齊用
                text.yview("end") #滾動到最下面
                text.config(state = 'disable')
                text.pack()
                notification.notify(title = '發行圖-新增', message = str((ans[2][15:])) ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )
                #data = {'message':ans}     # 設定要發送的訊息
                #data = requests.post(url, headers=headers, data=data)   # 使用 POST 方法
                with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'r', encoding = 'cp950') as f:
                    for txtlog in f:
                        log.append(str(txtlog))
                with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'w', encoding = 'cp950') as f:
                    for l in log:
                        l = l.replace('//', '\\\\')
                        l = l.replace('/', '\\')
                        f.write(str(l))
                        w_ans = str(ans[0]) + str(ans[1]) + str(ans[2]) + '\n'
                        w_ans = w_ans.replace('//', '\\\\')
                        w_ans = w_ans.replace('/', '\\')                    
                    f.write(w_ans)


        for dd in diff.files_deleted:
            log = []
            ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))," 圖面刪除:  ", dd
            #print(ans[0], Fore.RED + (ans[1]), Fore.CYAN + (ans[2][16:]))
            #notification.notify(title = '發行圖-刪除', message = str((ans[2][15:])) ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )
          # print("檔案-修改: ", diff.files_modified)
            text.config(state = 'normal')
            text.insert(tk.INSERT, str((ans[0][5:]) + ans[1] + ans[2][15:] + '\n'))
            text.tag_add('left', 1.0, "end")  #輸出文字對齊用
            text.yview("end") #滾動到最下面
            text.config(state = 'disable')
            text.pack()
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'r', encoding = 'cp950') as f:
                for txtlog in f:
                    log.append(str(txtlog))
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'w', encoding = 'cp950') as f:
                for l in log:
                    l = l.replace('//', '\\\\')
                    l = l.replace('/', '\\')
                    f.write(str(l))
                    w_ans = str(ans[0]) + str(ans[1]) + str(ans[2]) + '\n'
                    w_ans = w_ans.replace('//', '\\\\')
                    w_ans = w_ans.replace('/', '\\')                    
                f.write(w_ans)

            
        dm_ans = []
        for dm in diff.files_moved:
            log = []
            ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),' 變更或移動:   ' , dm[0],'   變更為   ', dm[1]
            #print(ans[0], Fore.MAGENTA + ans[1], Fore.CYAN + (ans[2][15:]), ans[3] + Fore.GREEN + (ans[4][15:]))
            #print(ans)
            text.config(state = 'normal')
            text.insert(tk.INSERT, str((ans[0][5:]) + ans[1] + ans[2][15:] + ans[3] + ans[4][15:] + '\n'))
            text.tag_add('left', 1.0, "end")  #輸出文字對齊用
            text.yview("end") #滾動到最下面
            text.config(state = 'disable')
            text.pack()
            dm_ans.append(dm)
            #notification.notify(title = '發行圖-變更檔名或移動', message = str(ans[3] + (ans[4][15:])) ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'r', encoding = 'cp950') as f:
                for txtlog in f:
                    log.append(str(txtlog))
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'w', encoding = 'cp950') as f:
                for l in log:
                    l = l.replace('//', '\\\\')
                    l = l.replace('/', '\\')
                    f.write(str(l))
                    w_ans = str(ans[0]) + str(ans[1]) + str((ans[2][15:])) + str(ans[3])  + str(ans[4]) + '\n'
                    w_ans = w_ans.replace('//', '\\\\')
                    w_ans = w_ans.replace('/', '\\')
                f.write(w_ans)



        dmf_ans = []
        for dmf in diff.files_modified:
            log = []
            #ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),' 圖面變更資訊:' , dm[0],' 變更為 ', dm[1]
            dmf_ans.append(dmf)
        if len(dmf_ans) < 1:
            pass
        else:
            #print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), Fore.GREEN +' 內容修改:檔案   ', dmf_ans[0],'  內容已修改')
            notification.notify(title = '發行圖-內容修改', message = str((dmf_ans[0][15:])) + '內容已修改' ,app_icon ='C:/Users/Public/Pictures/light.ico', timeout = 1 )
            text.config(state = 'normal')
            text.insert(tk.INSERT, str(time.strftime('%m-%d %H:%M:%S',time.localtime(time.time())) + ' 內容修改:  ' +  dmf_ans[0][15:] + '  內容已修改' + '\n'))
            text.tag_add('left', 1.0, "end")  #輸出文字對齊用
            text.yview("end") #滾動到最下面
            text.config(state = 'disable')
            text.pack()
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'r', encoding = 'cp950') as f:
                for txtlog in f:
                    log.append(str(txtlog))
            with open('C:/Users/Public/Documents/發行圖面更新紀錄.csv', 'w', encoding = 'cp950') as f:
                for l in log:
                    l = l.replace('//', '\\\\')
                    l = l.replace('/', '\\')
                    f.write(str(l))
                    w_ans = str(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))) + '內容修改: 檔案   ' + str((dmf_ans[0])) + '   內容已修改' + '\n'
                    w_ans = w_ans.replace('//', '\\\\')
                    w_ans = w_ans.replace('/', '\\')
                f.write(w_ans)




        # dirm_ans = [] 
        # for dirm in diff.d_modified:
        #     #ans = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))," 圖面移動: ", dirm
        #     dirm_ans.append(dirm)
        # if len(dirm_ans) < 2:
        #     pass
        # else:
        #     print('檔案從',dirm_ans[0],'    移至>>>    ',dirm_ans[1])
        #     #print(ans[0], ans[1], ans[2])
        #     #notification.notify(title = '發行圖-移動', message = dirm ,timeout = 0.3 )


        # print("資料夾-被修改: ", diff.dirs_modified)
        # print("資料夾-移動: ", diff.dirs_moved)
        # print("資料夾-刪除: ", diff.dirs_deleted)
        # print("資料夾-建立: ", diff.dirs_created)
		
        #pass
    
class DirMonitor(object):
    """文件夹监视类"""
    
    def __init__(self, aim_path):
        """构造函数"""
        
        self.aim_path= aim_path
        self.observer = Observer()
    
    def start(self):
        """启动"""
        
        event_handler = FileEventHandler(self.aim_path)
        self.observer.schedule(event_handler, self.aim_path, True)
        self.observer.start()
    
    def stop(self):
        """停止"""
        
        self.observer.stop()
    

 
if __name__ == "__main__":
    monitor = DirMonitor(r'//192.168.0.17/發行圖面/00-公旭圖面')
    monitor.start()
    monitor = DirMonitor(r'//192.168.0.17/發行圖面/00-群旭圖面')
    monitor.start()
    monitor = DirMonitor(r'//192.168.0.17/發行圖面/00-舊參考資料')
    monitor.start()
    monitor = DirMonitor(r'//192.168.0.17/發行圖面/01-編碼原則')
    monitor.start()

    win.protocol('WM_DELETE_WINDOW', on_exit)
    threading.Thread(target=icon.run, daemon=True).start()
    scrollbar.config(command=text.yview)

    win.mainloop()

    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        monitor.stop()
        monitor.join()






#def no_123():
    # print(Fore.CYAN +"{:=^80s}".format("歡迎使用_公旭群旭眾旭專用_發行圖面更新通知"))
    # time.sleep(1)
    # #授權
    # try:
    #     while True:
    #         ans = '5931'
    #         x = 3 #初始機會
    #         while x > 0 :
    #             x = x -1
    #             print('')
    #             pwd = input('請輸入登入密碼: ')
    #             if pwd == ans:
    #                 print('')
    #                 print('')
    #                 break
    #             else:
    #                 print('密碼錯誤!')
    #                 if x > 0:
    #                     print('還有', x,'次機會')
    #                 else:
    #                     print('輸入錯誤三次，程式結束')
    #                     print('')
    #                     print('')
    #                     print('3秒後程式關閉', end = '')
    #                     for i in range(6):
    #                         print("",end = ' ',flush = True)  #flush - 输出是否被缓存通常决定于 file，但如果 flush 关键字参数为 True，流会被强制刷新
    #                         time.sleep(0.5)
    #                         print('')
    #                     os._exit(0)                 
    #         break

    #     os.system('cls') #登入後清除畫面
    #     print('')
    #     print(Fore.CYAN +"{:=^80s}".format("歡迎使用_公旭群旭眾旭專用_發行圖面更新通知"))
    #     time.sleep(1)
    #     print('登入成功')
    #     print('')
    # except :
    #         print('Error :something went wrong')


    #line notification
    # url = 'https://notify-api.line.me/api/notify'
    # token = ''
    # headers = {'Authorization': 'Bearer ' + token}   # 設定權杖
