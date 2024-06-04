import PySimpleGUI as sg
import pandas as pd
import mysql.connector
import os

mysqldb=mysql.connector.connect(host="localhost",user="root",password="",database="aplikasi")
mycursor=mysqldb.cursor()


sg.theme('DarkGreen4')

EXCEL_FILE = 'Pendaftaran.xlsx'

df = pd.read_excel(EXCEL_FILE)

layout=[
[sg.Text('Masukan Data Kamu: ')],
[sg.Text('Nama',size=(15,1)), sg.InputText(key='Nama')],
[sg.Text('No Telp',size=(15,1)), sg.InputText(key='Tlp')],
[sg.Text('Alamat',size=(15,1)), sg.Multiline(key='Alamat')],
[sg.Text('Tgl Lahir',size=(15,1)), sg.InputText(key='Tgl Lahir'),
                                    sg.CalendarButton('Kalender', target='Tgl Lahir', format=('%Y-%m-%d'))],
[sg.Text('Jenis Kelamin',size=(15,1)), sg.Combo(['pria','wanita'],key='Jekel')],
[sg.Text('Hobi',size=(15,1)), sg.Checkbox('Belajar',key='Belajar'),
                            sg.Checkbox('Menonton',key='Menonton'),
                             sg.Checkbox('Musik',key='Musik')],
[sg.Submit(), sg.Button('clear'), sg.Button('view data'), sg.Button('open excel'), sg.Exit()]

]

window=sg.Window('Form pendaftaran',layout)

def select():
    results = []
    mycursor.execute("select nama,tlp,alamat,tgl_lahir,jekel,hobi from pendaftaran order by id desc")
    for res in mycursor:
        results.append(list(res))

    headings=['Nama','Tlp','Alamat','Tgl Lahir', 'Jekel', 'Hobi'] 

    layout2=[
        [sg.Table(values=results,
        headings=headings,
        max_col_width=35,
        auto_size_columns=True,
        display_row_numbers=True,
        justification='right',
        num_rows=20,
        key='-Table-',
        row_height=35)]
    ]   

    window=sg.Window("List Data", layout2)
    event, values = window.read()

def clear_input():
    for key in values:
        window[key]('')
        return None

while True :
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'Clear':
        clear_input()
    if event == 'view data':
        select()  
    if event == 'open excel':
        os.startfile(EXCEL_FILE)     
    if event == 'Submit':
        nama=values["Nama"]
        tlp=values["Tlp"]
        alamat=values["Alamat"]
        tgl_lahir=values["Tgl Lahir"]
        jekel=values["Jekel"]
        belajar=values["Belajar"]
        menonton=values["Menonton"]
        musik=values["Musik"]

        if belajar == True:
            hobi="Belajar"
        if menonton == True:
            hobi="Menonton"
        if musik == True:
            hobi="Musik"

        sql="insert into pendaftaran(nama,tlp,alamat,tgl_lahir,jekel,hobi) values(%s,%s,%s,%s,%s,%s)"
        val=(nama,tlp,alamat,tgl_lahir,jekel,hobi)
        mycursor.execute(sql,val)
        mysqldb.commit()
        
        df =df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Di Simpan')
        clear_input()
window.close()       
