import getpass
import os
import socket
from datetime import datetime
from uuid import getnode as get_mac
import pyautogui
import psutil
import platform
import win32com.client
import openpyxl
import cpuinfo   

name =str( getpass.getuser())    # Имя пользователя
ip =str( socket.gethostbyname(socket.getfqdn()))   # IP-адрес системы
mac = str(get_mac())   # MAC адрес
ost = str(platform.uname())    # Название операционной системы
cpu_info = str(cpuinfo.get_cpu_info()) #crjhjcnm hf,kns ghjwtccjhf
time=datetime.now()



Excel = win32com.client.Dispatch("Excel.Application") # создаем СОМ объект экселя, открываем книгу и делаем ее видимой
path = os.path.abspath('output.xlsx')
wb = Excel.Workbooks.Open(path) 
Excel.Visible=1
sheet=wb.ActiveSheet

data=[name,ip,mac,ost,cpu_info,time]

i=1
while i < 1000:
   # т.е. столбец постоянно 1, а строку мы ищем перебором
   val = sheet.Cells(i, 1).value
   if val == None:
       break
   i = i + 1
# когда мы нашли пустую строку
# нам в цикле нужно его заполнить
# данными из списка 
k = 1
for rec in data:
   sheet.Cells(i, k).value = rec
   k = k + 1
 
wb.Save()# сохраним
wb.Close()# и закроем
#Закроем COM объект
Excel.Quit()


workbook = openpyxl.load_workbook('output.xlsx')
sheet = workbook.active
# Получаем количество строк в листе
num_rows = sheet.max_row
# Получаем значение последней строки
last_row = [cell.value for cell in sheet[num_rows]]

last_name=str(last_row[0])#вытаскиввет необходмые параметры для сравнивания 
last_ost=str(last_row[3])
last_cpu_info=str(last_row[4])

incidients=open("Incidents.txt","a")#открывает файл для записи в него номера инцидента
if last_ost!=ost:#
   incidients.write("Подозрительное изменение опреационной системы ИНЦИДЕНТ1 ")
   incidients.write(str(time))
   incidients.write("\n")
if last_name!=name:
     incidients.write("Подозрительное изменение имени пользователя ИНЦИДЕНТ2 ")
     incidients.write(str(time))
     incidients.write("\n")
if last_cpu_info!=cpu_info:
     incidients.write("Подозрительное изменение ифнормации о процессоре ИНЦИДЕНТ3 ")
     incidients.write(str(time))
     incidients.write("\n")
cpu_percent = psutil.cpu_percent(interval=1) 
if cpu_percent>=80:
     incidients.write("Подозрительная нагрузка процессора ИНЦИДЕНТ4 ")
     incidients.write(str(time))
     incidients.write("\n")
incidients.close()