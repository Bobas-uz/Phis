import csv
import pandas as pd
import xlrd
from tkinter import Tk     
from tkinter.filedialog import askopenfilename
import sys
from PyQt5 import QtWidgets
import dialog


#excel file choose
def get_files():
    Tk().withdraw() 
    #відкриття файлу через Qt
    #filename = QtWidgets.QFileDialog.getOpenFileName(self,tr("Виберіть файл виписки"),tr("Excel files(*.xls,*.xlsx,*.csv)"))
    filename = askopenfilename( title="Виберіть файл виписки",filetypes=[("Excel files", ".xlsx .xls")])
    return filename 
#file справочник choose
def get_dict():
    Tk().withdraw() 
    filename = askopenfilename( title="Виберіть файл довіника",filetypes=[("Excel files", ".xlsx .xls ")])
    return filename 
#clear input file виписка
def clear_input(path):
    wb = xlrd.open_workbook(path, encoding_override='cp1251')
    ws = wb.sheet_by_index(0)
    data=[]
    dict = get_dict()
    wb1 = xlrd.open_workbook(dict, encoding_override='cp1251')
    ws1 = wb1.sheet_by_index(0)
    kontragent = {}
    for row1 in range(ws1.nrows):
        #скорочена версія довідника str(ws1.cell_value(row1,2))** повний довідник str(ws1.cell_value(row1,12))
        kontragent[str(ws1.cell_value(row1,1))] = str(ws1.cell_value(row1,2))
   
    out_name = str(ws.cell_value(4, 2))+'.csv'
    with open(out_name,'w', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        for row in range(5,ws.nrows):
            
            data.append(str(kontragent.get(str(ws.cell_value(row, 6)))))
            data.append(str(ws.cell_value(row, 6)))
            data.append(ws.cell_value(4, 2))
            data.append(str(ws.cell_value(row, 2)))
            data.append(str(ws.cell_value(row, 3)).replace('.',','))
            data.append('40314229')
            data.append('UA818201720313291001201094235')
            if (str(ws.cell_value(row, 11)) == '6.7.5'):
                writer.writerow(data)
            data=[]
class ExampleApp(QtWidgets.QMainWindow, dialog.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
def main():
    #app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
   # window = ExampleApp()  # Создаём объект класса ExampleApp
    #window.show()  # Показываем окно
    #app.exec_()  # и запускаем приложение
    filename = get_files()
    clear_input(filename)
    

if __name__=='__main__':
    main()
    
    


 


base = pd.read_excel("Справочник.xls")

