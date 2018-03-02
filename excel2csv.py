###################################
#      Rico FAUCHARD - 2018       #
#      (conversion excel en csv)  #
###################################

from Tkinter import Frame, Tk, BOTH, Text, Menu, END
import tkFileDialog
import xlrd
import csv

class UI(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent)   

        self.parent = parent        
        self.initUI()

    def initUI(self):

        self.parent.title("Convertir Excel en CSV")
        self.pack(fill=BOTH, expand=1)

        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        fileMenu = Menu(menubar)
        
        fileMenu.add_command(label="Ouvrir", command=self.onOpen)
        menubar.add_cascade(label="Fichier", menu=fileMenu)        

        self.txt = Text(self)
        self.txt.pack(fill=BOTH, expand=1)


    def onOpen(self):

        ftypes = [('Excel', '*.xlsx'), ('Tous les fichiers', '*')]
        dlg = tkFileDialog.Open(self, filetypes = ftypes)
        fl = dlg.show()

        if fl != '':
            self.excel_to_csv(fl)
            self.txt.insert(END, "Converion ok")

    def readFile(self, filename):
        f = open(filename, "r")
        text = f.read()
        return text
    
    
    def excel_to_csv(self,excel_path):
        csv_path = excel_path.replace('xlsx','csv')
        workbook = xlrd.open_workbook(excel_path)
        worksheet = workbook.sheet_by_index(0)
        csvfile = open(csv_path, 'wb')
        wr = csv.writer(csvfile, delimiter=';',quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for rownum in xrange(worksheet.nrows):
            wr.writerow(
                list(x.encode('utf-8') if type(x) == type(u'') else x
                      for x in worksheet.row_values(rownum)))
        csvfile.close()
         
def main():
    root = Tk()
    ex = UI(root)
    root.geometry("300x250+300+300")
    root.mainloop()  

if __name__ == '__main__':
    main()
