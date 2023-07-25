import time
import webbrowser as web

import pygal
from pynput.keyboard import GlobalHotKeys, HotKey, Key, Listener

count ={}
keys =[]
keyin = []
print('开始统计键盘输入，热键：Ctrl+t结束统计并生成报表')

def on_press(key):
    if str(key) == r"'\x14'":
        return False
    string = str(key)
    keyin.append(string) 
     
    if key in count:
        count[key] += 1
    else:
        count[key] = 1
    print(count)  
with Listener(on_press=on_press) as listener:
    listener.join()

with open('keys.txt', 'a') as f:
    nowtime = str(time.strftime("%Y-%m-%d %H:%M:%S",time.localtime()))
    f.write(nowtime+':')   
    f.write(str(keyin)+'\n')   
f.close()
hist = pygal.Bar()
hist.title ='keyboard按键输入频率统计'
# hist.x_title =  "Key"
hist.y_title = "Frequency of Key"
for key in count:
    keys.append(key)
    hist.add('{0}'.format(key),count[key])
# hist.x_labels = keys
hist.render_to_file('./xxx.svg')
# web.get('chrome').open_new_tab('./xxx.svg') 
print("生成统计与日志文件完成...")

