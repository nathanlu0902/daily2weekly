from tkinter import *
from tkinter.filedialog import askopenfilename

class inputFile:

    def __init__(self,master):
        self.inputFile=Frame(master=master,bd=2,relief="sunken")
        self.inputFile.grid(row=0,column=0,padx=2, pady=2)
        self.add_btn()
        self.add_path_label('Input File Path')

    def add_path_label(self,text):
        self.label=Label(text=text,width=200,height=10,master=self.inputFile)
        self.label.grid(column=0,row=0)


    def add_btn(self):
        self.btn=Button(text='Choose Daily Template',master=self.inputFile,command=self.chooseFile)
        self.btn.grid(column=0,row=1)


    def chooseFile(self):
        filePath=askopenfilename(filetypes=[('Excel xlsx Files', '*.xlsx')])
        if filePath:
            self.label.config(text=filePath)


class covert:
    def __init__(self,master):
        self.output=Frame(master=master,bd=2,relief='sunken')
        self.output.grid(row=1,column=0,padx=2,pady=2)
        self.cutOffDaySel()
        self.add_btn()

    def show_day_label(self,text):
        self.day_label=Label(text=text,master=self.output)
        self.day_label.grid(row=0,column=0)


    def cutOffDaySel(self):
        var = StringVar()
        r1 = Radiobutton(self.output, text='Monday', variable=var, value='Mon', command=self.show_day_label(var.get()))
        r1.grid(row=1,column=0)
        r2 = Radiobutton(self.output, text='Tuesday', variable=var, value='Tue', command=self.show_day_label(var.get()))
        r2.grid(row=1,column=1)
        r3 = Radiobutton(self.output, text='Wednesday', variable=var, value='Wed', command=self.show_day_label(var.get()))
        r3.grid(row=1,column=2)
        r4 = Radiobutton(self.output, text='Thursday', variable=var, value='Thu', command=self.show_day_label(var.get()))
        r4.grid(row=1,column=3)
        r5 = Radiobutton(self.output, text='Friday', variable=var, value='Fri', command=self.show_day_label(var.get()))
        r5.grid(row=1,column=4)
        r6 = Radiobutton(self.output, text='Satday', variable=var, value='Sat', command=self.show_day_label(var.get()))
        r6.grid(row=1,column=5)


    def add_btn(self):
        btn=Button(master=self.output,text='Convert to Weekly',command=)




if __name__=='__main__':

    root=Tk()
    root.title('Daily LTD SP to Weekly')
    root.rowconfigure(1,minsize=40,weight=1)
    root.columnconfigure(2,minsize=40,weight=1)
    inputFile=inputFile(root)
    covert=covert(root)

    root.mainloop()