import os
import sys
import time

import pythoncom
import win32api
import win32com.client
import win32con
import win32gui
from PIL import ImageGrab
from WorkWeixinRobot.work_weixin_robot import WWXRobot


def check_dirpath():
    # 在没有png目录时，创建一个png目录
    #os.chdir(os.path.dirname(__file__))   可能导致程序打包运行路径为系统temp路径而非程序路径而找不到文件夹
    os.chdir(os.path.dirname(os.path.realpath(sys.argv[0]))) 
    print("当前路径"+os.getcwd())
    absPath = os.path.abspath('img')
    os.path.dirname(os.path.realpath(sys.argv[0]))
    path = [x for x in os.listdir('img') if os.path.isdir(os.path.join(absPath,x))] 
    print("检测图片保存目录png是否存在...")
    if 'png' in path:
        print('已存在png目录\n')
        pass
    else:
        print('png目录不存在，创建png目录')
        #创建目录
        pngPath = os.path.join(absPath,'png')
        os.mkdir(pngPath) 
        print('目录创建成功\n')  

def Screening(coordinate):
    #截屏发送API
    nowtime = time.strftime('%Y_%m_%d_%H_%M_%S',time.localtime(time.time()))
    print('截图,保存中...')
    time.sleep(1)
    im = ImageGrab.grab(bbox=coordinate)  #bbox=(441,149,1478,926)
    #im.show()
    #im.save(r'png\%s.png' %(nowtime))
    im.save("img\png\%s.png" %(nowtime))
    print('%s.png保存成功\n' %(nowtime))
    #发送企业微信
    # wwx = WWXRobot(key='ed096bb8-3a5f-4326-b743-5ba009bba2af')
    # wwx.send_image(local_file='img\png\%s.png' %(nowtime))
    # print("发送企业微信API完成\n")

def truncate_file(time_tmp):
    """按时间删除文件和清空文件
    :param path 目录
    :param time_tmp 多少天前的文件 int
    """
    print('开始清理1天前截图...')
    dir_path = os.path.dirname(os.path.realpath(sys.argv[0]))
    print('dir_path:'+ dir_path)
    pngdir_path=os.path.join(dir_path,'img','png')
    print('pngdir:'+ pngdir_path)
    print(os.path.isdir(pngdir_path))
    if not os.path.isdir(pngdir_path):
        print("目录不存在,程序退出.....")
        sys.exit()
    # 过去的时间
    past_time = int(time.time()) - (3600 * 24 * time_tmp)
    os.chdir(pngdir_path)
    filelist = os.listdir(pngdir_path)
    #print(filelist)
    if filelist is None:
        print("没有文件，不做操作")
        sys.exit(0)
    #遍历文件
    for file in filelist:
        # 获取文件修改时间
        modify_time = os.path.getmtime(file)
        #print(modify_time)
        filename = pngdir_path + "/" + file
        #rocketmq_client.log 开头的清空
        if modify_time < past_time:
            print("正在删除文件:" + filename)
            os.remove(file)
        else:
            continue
    print('清理完成\n')


def get_window_pos(name):
    name = name
    handle = win32gui.FindWindow('GlassWndClass-GlassWindowClass-2', name)#
    # 发送还原最小化窗口的信息
    win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
    # 设为高亮
    pythoncom.CoInitialize()
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    # win32api.keybd_event(13, 0, 0, 0)
    win32gui.SetForegroundWindow(handle)
    # 获取窗口句柄
    if handle == 0:
        return None
    else:
        # 返回坐标值
        return win32gui.GetWindowRect(handle)

pic_savedate = 1   #截图保留日期 单位(day)
sleep_time = 60   #单位（s）

while True:

    coordinate = get_window_pos('SOPAS Engineering Tool 2022.1') 
    check_dirpath()
    Screening(coordinate)
    truncate_file(pic_savedate)
    os.chdir("D:\sickListen")
    print("休眠，"+ str(sleep_time) +"s后开始下一次截图发送...\n")  
    time.sleep(sleep_time) 



