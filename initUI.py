import random,psycopg2
from openpyxl import Workbook
from calendar import monthrange
from datetime import datetime
from tkinter import *
from API import Shift

class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.root = parent
        self.initUI()
        self.arr = []
        self.banned_days = {}
        self.conn = psycopg2.connect(host="localhost", database="excelSheet", user="postgres", password="root")
        self.cursor = self.conn.cursor()
        self.shift = Shift()

    def initUI(self):
        #Bu kisimda genel olarak arayuz yapilmaktadir.
        self.grid()
        cizgi = Canvas(self,height=1,width=500)
        cizgi.grid(row=0,column=0,columnspan=10)

        self.dayList = Listbox(self,height=7,width=15)
        for i in ['Pazartesi','Salı','Çarşamba','Perşembe','Cuma','Cumartesi','Pazar']:
            self.dayList.insert(END,i)
        self.dayList.grid(row=1,column=1,columnspan=3,rowspan=7)
        label = Label(self,text="Birden fazla kişi girebilirsiniz\nİki kişi arasını virgül(,) ile ayırınız.")
        label.grid(row=2,column=5,columnspan=6)
        self.persons = Entry(self)
        self.persons.grid(row=3,column=6,columnspan=3)

        offDay = Button(self,text="İzin Ekle",command=self.addOffDay)
        offDay.grid(row=4,column=6,columnspan=3)

        cizgi = Canvas(self,height=1,width=500,background="black")
        cizgi.grid(row=9,column=0,columnspan=10)

        label = Label(self, text="Çalışan ekle/çıkar")
        label.grid(row=10,column=3,columnspan=3)

        self.addRemoveText = Entry(self,width=30)
        self.addRemoveText.grid(row=11, column=2,columnspan=5)



        add = Button(self,text="Ekle",command=self.addPerson)
        add.grid(row=11,column=1)
        remove = Button(self,text="Çıkar",command=self.deletePerson)
        remove.grid(row=12,column=1)

        start = Button(self,text="Başlat",command=self.start)
        start.grid(row=13,columnspan=10)

    def addPerson(self):
        name = self.addRemoveText.get()
        self.shift.addPerson(name)

    def deletePerson(self):
        name = self.addRemoveText.get()
        self.shift.removePerson(name)

    def addOffDay(self):
        day = self.dayList.get(self.dayList.curselection()[0])
        persons = self.persons.get().split(",")
        for persoName in persons:
            self.shift.offDay(persoName, day)

    def start(self):
        year = 2021
        month = 2
        self.shift.parseMonth(year, month)
        self.shift.createExcel()
        self.shift.personelPut("WORK1")
        self.shift.personelPut("WORK2",offset=1)
        self.shift.offset()



def main():
    root = Tk()
    root.title("Shift Control")
    root.resizable(0,0)
    root.geometry("500x350+1000+500")
    App = Example(root)
    root.mainloop()
if __name__ == '__main__':
    main()
