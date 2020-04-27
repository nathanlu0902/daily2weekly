from tkinter import *
from tkinter.filedialog import askopenfilename
from app import Excel_handler
import tkinter.messagebox as mb
import os


inputFilePath=''


class Input:

    def __init__(self,master):
        self.master=master
        self.add_btn()
        self.add_path_label()

    def add_path_label(self):
        self.label=Label(text=inputFilePath,master=self.master,borderwidth=2,relief='flat')
        self.label.grid(column=1,row=1,columnspan=2,pady=10)


    def add_btn(self):
        self.btn=Button(text='Choose Daily Template',master=self.master,command=lambda:self.chooseFile())
        self.btn.grid(column=1,row=2,columnspan=2,pady=10)


    def chooseFile(self):
        global inputFilePath

        inputFilePath=askopenfilename(filetypes=[('Excel xlsx Files', '*.xlsx')])

        if os.path.splitext(inputFilePath)[-1]!='.xlsx':
            mb.showerror('格式错误','请选择Excel文件')
            return

        self.label.config(text=inputFilePath)



class Output:

    def __init__(self,master):
        self.master=master
        self.cutOffDaySel()
        self.add_btn()
        self.add_label()


    def add_label(self):
        self.label=Label(text='',master=self.master)
        self.label.grid(row=4,column=1)


    def cutOffDaySel(self):

        frame=Frame(self.master,relief='sunken',borderwidth=5)
        frame.grid(row=3,column=1,columnspan=2,pady=25)

        var = StringVar()
        r1 = Radiobutton(frame, text='Monday', variable=var, value='Mon', command=lambda:self.change_day_label(text=var.get()))
        r1.grid(row=0,column=0)
        r2 = Radiobutton(frame, text='Tuesday', variable=var, value='Tue', command=lambda:self.change_day_label(text=var.get()))
        r2.grid(row=0,column=1)
        r3 = Radiobutton(frame, text='Wednesday', variable=var, value='Wed', command=lambda:self.change_day_label(text=var.get()))
        r3.grid(row=0,column=2)
        r4 = Radiobutton(frame, text='Thursday', variable=var, value='Thu', command=lambda:self.change_day_label(text=var.get()))
        r4.grid(row=1,column=0)
        r5 = Radiobutton(frame, text='Friday', variable=var, value='Fri', command=lambda:self.change_day_label(text=var.get()))
        r5.grid(row=1,column=1)
        r6 = Radiobutton(frame, text='Saturday', variable=var, value='Sat', command=lambda:self.change_day_label(text=var.get()))
        r6.grid(row=1,column=2)
        r7 = Radiobutton(frame, text='Sunday', variable=var, value='Sun', command=lambda:self.change_day_label(text=var.get()))
        r7.grid(row=1,column=3)


    def change_day_label(self,text):
        self.cutOffDay = text
        self.label.config(text='你选择每周按%s计算weekly SP'%text)


    def add_btn(self):
        btn=Button(master=self.master,text='Convert to Weekly',command=lambda:self.start_convert())
        btn.grid(row=4,column=2,padx=25,pady=25)

    def gen_output_file_path(self):
        directory = os.path.split(inputFilePath)[0]
        filename = os.path.split(inputFilePath)[-1]
        out_filename = os.path.splitext(filename)[0] + '-weekly.xlsx'
        outputFilePath = os.path.join(directory, out_filename)

        return outputFilePath

    def successful(self):
        mb.showinfo('成功','weekly sp已生成于%s'%self.gen_output_file_path())


    def start_convert(self):

        input=inputFilePath
        output=self.gen_output_file_path()

        print(input,output)

        if input is None:
            mb.showerror('未选择模板', '请选择weekly模板')

        if self.cutOffDay is None:
            mb.showerror('未选择日期', '请选择Cut Off Day')

        else:
            handler=Excel_handler(input,self.cutOffDay)
            handler.write_to_excel(output)
            self.successful()






if __name__=='__main__':

    root=Tk()
    root.title('Daily LTD SP to Weekly')

    content=Frame(root)
    # frame=Frame(content,borderwidth=5,relief='sunken',width=300,height=300)
    content.grid(column=0,row=0)
    # frame.grid(column=0,row=0,columnspan=3,rowspan=4)


    input=Input(content)
    output = Output(content)

    root.rowconfigure(0,weight=1)
    root.columnconfigure(0, weight=1)
    content.rowconfigure(0,weight=1)
    content.rowconfigure(1, weight=1)
    content.rowconfigure(2, weight=1)
    content.rowconfigure(3, weight=1)
    content.rowconfigure(4, weight=1)
    content.columnconfigure(0,weight=1)
    content.columnconfigure(1, weight=1)
    content.columnconfigure(2, weight=1)
    content.columnconfigure(3, weight=1)


    root.mainloop()