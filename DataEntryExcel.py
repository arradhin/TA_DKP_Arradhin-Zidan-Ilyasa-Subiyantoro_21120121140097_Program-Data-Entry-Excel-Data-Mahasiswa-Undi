import PySimpleGUI as sg
import pandas as pd # untuk menyambungkan python ke excel

# tema 
sg.theme ('GrayGrayGray')

# variabel untuk mengakses file excel 
Excel_path = "C:\\Users\\Arradhin\\Documents\\Data_Entry.xlsx"
df = pd.read_excel(Excel_path) #dataframe

class MyClass:
  def __init__(self) -> True:
      pass

#elemen
layout = [

    [sg.Text("Program Data Entry Excel")],
    [sg.Text("Nama" , size=(15,1)), sg.InputText(key = "Nama")], # nama heading di excel harus sesuai dengan key
    [sg.Text("NIM" , size=(15,1)), sg.InputText(key = "NIM")],
    [sg.Text("Jenis kelamin", size=(15,1)),sg.Combo(["Laki-laki", "Perempuan"], key = "Jenis kelamin")],
    [sg.Text("Asal kota" , size=(15,1)), sg.InputText(key = "Asal kota")],
    [sg.Text("Angkatan", size=(15,1)),sg.Combo(["2017","2018","2019","2020","2021"], key = "Angkatan")],
    [sg.Text("Semester", size=(15,1)), sg.Spin([i for i in range(1,9)],
        initial_value=1, key= "Semester")],
    [sg.Submit(),sg.Button("Clear"), sg.Exit()]

]

# judul window
window = sg.Window("Data Entry", layout)

# clear form
def clear_input(): #method
    for key in values:
        window[key]('')
    return None

#loop
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit": # menutup aplikasi jika mengklik tombol X atau tombol exit
        break 
    if event == "Clear":
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index= True)
        df.to_excel(Excel_path, index=False)
        sg.popup("Data Tersimpan") # popup setelah submit
window.close()
