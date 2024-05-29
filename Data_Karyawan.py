import PySimpleGUI as sg
import pandas as pd

# Tentukan tema (opsional)"
sg.theme('DarkGreen4')

# Pilih lokasi penyimpanan file Excel"
EXCEL_FILE = 'Pendaftaran.xlsx'

# Membaca data yang ada (jika file tersedia)
try:
    df = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nama', 'Nik', 'Tlp', 'Alamat', 'Tgl Lahir',
     'Jekel', 'Aluminium', 'Fo', 'Foc', 'Office', 'Packing', 'Warehouse'])

# Mendefinisikan elemen tata letak
layout =[
    [sg.Text('Masukan Data Kamu:')],
    [sg.Text('Nama', size=(15, 1)), sg.InputText(key='Nama')],
    [sg.Text('Nik', size=(15, 1)), sg.InputText(key='Nik')],
    [sg.Text('No Telp', size=(15, 1)), sg.InputText(key='Tlp')],
    [sg.Text('Alamat', size=(15, 1)), sg.Multiline(key='Alamat')],
    [sg.Text('Tgl Lahir', size=(15, 1)), sg.InputText(key='Tgl Lahir'),
     sg.CalendarButton('Kalender', target='Tgl Lahir', format='%d-%M-%Y')],
    [sg.Text('Jenis Kelamin', size=(15, 1)), sg.Combo(['pria', 'wanita'], key='Jekel')],
    [sg.Text('Departement', size=(15, 1)),
     sg.Checkbox('Aluminium', key='Aluminium'),
     sg.Checkbox('Fo', key='Fo'),
     sg.Checkbox('Foc', key='Foc'),
     sg.Checkbox('Office', key='Office'),
     sg.Checkbox('Packing', key='Packing'),
     sg.Checkbox('Warehouse', key='Warehouse')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()],
]


# Buat jendela
window = sg.Window('Form Data Karyawan PT.ZTT Cable Indonesia', layout)

def clear_input():
    """Membersihkan semua kolom input."""
    for key in window.keys():
        if isinstance(window[key], (sg.InputText, sg.Multiline)):
            window[key]('') 

# Setel nilai ke string kosong
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
      # Tambahkan data baru ke DataFrame
        df = df.append (values, ignore_index=True)

        # Coba simpan ke Excel, tangani potensi kesalahan.
        try:
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data Berhasil Di Simpan')
        except PermissionError:
            sg.popup('Error: Insufficient permissions to save file. Try running with administrator privileges.')
        except Exception as e:
            sg.popup(f'Error saving data: {e}')

        # Bersihkan kolom input setelah data terkirim dengan sukses.
        clear_input()

window.close()
