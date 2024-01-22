import tkinter as tk
import time
from tkinter import filedialog
from tkinter import scrolledtext
from openpyxl import Workbook
import serial
from csv import reader
from io import StringIO

win = tk.Tk()
win.title('Temp')
win.geometry('800x600')
win['bg'] = 'gray'

ser = serial.Serial("COM3", 115200, timeout=1)

text_widget1 = scrolledtext.ScrolledText(win, state ='disable', width=42, height=19, wrap=tk.WORD)
text_widget1.place(x=10, y=40)
text_widget3 = scrolledtext.ScrolledText(win, state ='disable', width=42, height=19, wrap=tk.WORD)
text_widget3.place(x=430, y=40)
text_widget2 = scrolledtext.ScrolledText(win, state ='disable',width=95, height=5, wrap=tk.WORD)
text_widget2.place(x=10, y=450)

entry1 = tk.Entry(win, width=20,font= ('Arial',15))
entry1.place(x= 380, y=540)
entry2 = tk.Entry(win, width=5,font= ('Arial',15))
entry2.place(x=180 , y=540)
          
def export_to_excel(node):
    if (node==1):
        content = text_widget1.get("1.0", tk.END)
    else:
        if (node==2):
            content = text_widget3.get("1.0", tk.END)
        else:
            return 0
    workbook = Workbook()
    sheet = workbook.active
    csv_data = reader(StringIO(content))
    for row_num, row_data in enumerate(csv_data, start=1):
        for col_num, cell_data in enumerate(row_data, start=1):
            sheet.cell(row=row_num, column=col_num, value=cell_data)
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        workbook.save(file_path)
    return 1

def check_hex(hex):
    try:
        for i in hex:
            int(i, 16)
            return True
    except ValueError:
        return False
       
def read_com():
    receive = ser.readline().decode('utf-8').strip()
    if(receive):
        if(check_hex(receive)):
            sum =0
            for i in range(6):
                sum+= int(receive[i], 16)
            if(sum == (int(receive[6],16)*16+int(receive[7],16))):
                node = int(receive[0])*16+int(receive[1],16)
                leng = int(receive[2])*16+int(receive[3],16)
                temp = int(receive[4])*16+int(receive[5],16)
                return node,temp
        else: return 0
    else:
        current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, f"{current_time}"+"    Khong nhan duoc\n")
        text_widget2.configure(state="disable")
        text_widget2.see(tk.END)
        return 0

def display(widget,data):
    text_widget1.configure(state="normal")
    text_widget3.configure(state="normal")
    current_time = time.strftime("%d-%m-%Y , %H:%M:%S")
    if(widget == 1):
        text_widget1.insert(tk.END, f"{current_time} "+", "+str(data) +"\n")
        text_widget1.see(tk.END)
    else: 
        text_widget3.insert(tk.END, f"{current_time} "+", "+str(data) +"\n")
        text_widget3.see(tk.END)
    text_widget1.configure(state="disable")
    text_widget3.configure(state="disable")
    
def update_data():
    receive = ser.readline().decode('utf-8').strip() 
    if(receive):
        #sum =0
        # for i in range(6):
        #     sum+= int(receive[i], 16)
        # if(sum == (int(receive[6],16)*16+int(receive[7],16))):
        temp = (int(receive[1])*1000+int(receive[2])*100+int(receive[3])*10+int(receive[4]))/100
        display(int(receive[0]),temp)
        # current_temp  = tk.Button(win, text= 'Current temp: '+{temp}, width = 10, height=1)
        # current_temp.place(x=180, y= 350)
        # else:
        # current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        # text_widget2.configure(state="normal")
        # text_widget2.insert(tk.END, f"{current_time}"+"    Sai checksum\n")
        # text_widget2.configure(state="disable")
        # text_widget2.see(tk.END)
        win.after(2000, update_data)  
    else:
        win.after(1000, update_data)
        # current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        # text_widget2.configure(state="normal")
        # text_widget2.insert(tk.END, f"{current_time}"+"    Khong nhan duoc\n")
        # text_widget2.configure(state="disable")
        # text_widget2.see(tk.END)
update_data()

def click_excel():
    num = int(entry2.get())
    if(export_to_excel(num)):
        current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, f"{current_time}"+"    Xuat file Excel\n")
        text_widget2.configure(state="disable")
        text_widget2.see(tk.END)
    else:
        current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, f"{current_time}"+"    Vui long nhap dung so Node\n")
        text_widget2.configure(state="disable")
    entry2.delete(0, 'end')
but = tk.Button(win, text= 'Excel', width = 10, height=1,command = click_excel)
but.place(x=180, y=570)

def send_data(node,warn):
    #hexa = hex(warn).lstrip('0x').zfill(2)
    data = str(warn)
    # checksum = 0
    # for i in range(6):657
    #     checksum+= int(data[i], 16)
    data_send= str(warn)#data+hex(checksum).lstrip('0x').zfill(2)
    try:
        ser.write(str(data).encode())
    except Exception as e:
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, 'Loi khi gui du lieu: '+ str(e)+'\n')
        text_widget2.configure(state="disable")

def click_send():
    node = (entry2.get())
    warn = (entry1.get())
    if ((node!="")&(warn!="")):
        current_time = time.strftime("%d-%m-%Y %H:%M:%S")
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, f"{current_time}"+ " Node "+node+" Gui nguong canh bao: "+ warn+" do\n")
        text_widget2.configure(state="disable")
        text_widget2.see(tk.END)
        entry1.delete(0, 'end')
        entry2.delete(0, 'end')
        send_data(node,warn)
    else:
        text_widget2.configure(state="normal")
        text_widget2.insert(tk.END, "Vui long nhap nhiet do can gui\n")
        text_widget2.configure(state="disable")
        text_widget2.see(tk.END)
but = tk.Button(win, text= 'Send', width = 10, height=1,command = click_send)
but.place(x=380, y=570)
status = tk.Label(win, text = 'Status', font=14)
status.place(x=10, y= 425)
tk.Label(win,text='Node 1',font=14).place(x=140,y=10)
tk.Label(win,text='Node 2', font=14).place(x=580,y=10)
win.mainloop()


    