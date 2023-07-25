import pandas as pd
from pynput.keyboard import Key, Listener
import tkinter as tk
from tkinter import filedialog
import os


def choise_excel():
        # root = tk.Tk()
        # root.withdraw()
        path = filedialog.askopenfilename()  #获取文件夹中的某文件
        df = pd.read_excel(path)
        col1 = df.columns.to_list()
        for numb in range(0,len(col1)):
                print(str(numb)+':'+col1[numb])
        return col1,path

def ckeck():
        path = list(os.listdir('E:\\'))
        if 'CreateExcel' in path:
                print('CreateExcel已存在。')
                pass
        else:
                print('CreateExcel不存在，创建CreateExcel。')
                #创建目录
                os.mkdir('E:\CreateExcel')


def Create(path,column):

    status = True
    try:
        df = pd.read_excel(path, index_col=0)
    except FileNotFoundError as e:
        print('路径下找不到文件，请确认'+path+'文件存在！\n', e)
        status = False

    while status:
        try:
                for value in df[column].unique():
                        df_temp = df[df[column] == value]
                        
                        df_temp.to_excel("E:\CreateExcel\%s.xlsx" % (value))
                        print(value + '创建完成!')
                print('已完成所有excel创建！')
                status = False
        except:
                print('无法拆分，请确认所选行的数据类型正确！')
                break


def on_press(key):
    if key == Key.esc:
        return False

def input_number(col1): 
  staus =True      
  while staus:

        try:
                num = int(input('请按对应数字选择要拆分的列:'))
                if(num in list(range(0,len(col1)))):
                        staus=False
                        return num                        
                else:
                        print('请输入有效的数字！')
                        continue 
        except:
                print('请输入一个数字类型！')
                continue 




col1,path=choise_excel()
index = input_number(col1)
os.system('cls')
column = col1[index]

ckeck()
Create(path,column)
print('按ESC键退出序~')
with Listener(on_press=on_press) as listener:
    listener.join()

print('关闭程序')


