from tkinter import *
from tkinter.filedialog import askopenfilename
from app import Excel_handler
import tkinter.messagebox as mb
import os


inputFilePath = ''


class Input:

    def __init__(self, master):
        self.master = master
        self.add_guide_label()
        self.add_btn()
        self.add_path_label()

    def add_guide_label(self):
        self.label = Label(
            text='第一步：选择LTD Daily模板',
            master=self.master,
            borderwidth=2,
            relief='flat')
        self.label.grid(column=0, row=0, padx=15, columnspan=4, sticky=W)

    def add_path_label(self):
        self.label = Label(
            text=inputFilePath,
            master=self.master,
            borderwidth=2,
            relief='flat',
            wraplength=400,
            justify='left')
        self.label.grid(column=0, row=1, padx=15, columnspan=4,rowspan=2, pady=5)

    def add_btn(self):
        self.btn = Button(
            text='Choose Daily Template',
            master=self.master,
            command=lambda: self.chooseFile())
        self.btn.grid(column=0, row=3, columnspan=2, padx=15, sticky=W)

    def chooseFile(self):
        global inputFilePath

        inputFilePath = askopenfilename(
            filetypes=[('Excel xlsx Files', '*.xlsx')])

        if os.path.splitext(inputFilePath)[-1] != '.xlsx':
            mb.showerror('格式错误', '请选择Excel文件')
            return

        self.label.config(text=inputFilePath)


class Output:

    def __init__(self, master):
        self.master = master
        self.add_guide_label()
        self.cutOffDaySel()
        self.add_btn()
        self.add_label()

    def add_guide_label(self):
        self.label = Label(
            text='第二步：选择每周Cutoff日期',
            master=self.master,
            borderwidth=2,
            relief='flat')
        self.label.grid(
            column=0,
            row=4,
            padx=15,
            columnspan=3,
            pady=5,
            sticky=W)

    def cutOffDaySel(self):

        frame = Frame(self.master, relief='sunken', borderwidth=5)
        frame.grid(row=5, column=0, columnspan=3, padx=15, pady=5, sticky=W)

        var = StringVar()
        r1 = Radiobutton(
            frame,
            text='Monday',
            variable=var,
            value='Mon',
            command=lambda: self.change_day_label(
                text=var.get()))
        r1.grid(row=0, column=0)
        r2 = Radiobutton(
            frame,
            text='Tuesday',
            variable=var,
            value='Tue',
            command=lambda: self.change_day_label(
                text=var.get()))
        r2.grid(row=0, column=1)
        r3 = Radiobutton(
            frame,
            text='Wednesday',
            variable=var,
            value='Wed',
            command=lambda: self.change_day_label(
                text=var.get()))
        r3.grid(row=0, column=2)
        r4 = Radiobutton(
            frame,
            text='Thursday',
            variable=var,
            value='Thu',
            command=lambda: self.change_day_label(
                text=var.get()))
        r4.grid(row=1, column=0)
        r5 = Radiobutton(
            frame,
            text='Friday',
            variable=var,
            value='Fri',
            command=lambda: self.change_day_label(
                text=var.get()))
        r5.grid(row=1, column=1)
        r6 = Radiobutton(
            frame,
            text='Saturday',
            variable=var,
            value='Sat',
            command=lambda: self.change_day_label(
                text=var.get()))
        r6.grid(row=1, column=2)
        r7 = Radiobutton(
            frame,
            text='Sunday',
            variable=var,
            value='Sun',
            command=lambda: self.change_day_label(
                text=var.get()))
        r7.grid(row=1, column=3)

    def add_label(self):
        self.label = Label(text='', master=self.master, relief='flat')
        self.label.grid(row=6, column=0, padx=15, pady=5, sticky=W)

    def change_day_label(self, text):
        self.cutOffDay = text
        self.label.config(text='你已选择每周卡%s计算Weekly交期' % text)

    def add_btn(self):
        btn = Button(
            master=self.master,
            text='Convert to Weekly',
            command=lambda: self.start_convert())
        btn.grid(row=7, column=0, padx=15, sticky=W)

    def gen_output_file_path(self):
        directory = os.path.split(inputFilePath)[0]
        filename = os.path.split(inputFilePath)[-1]
        out_filename = os.path.splitext(filename)[0] + '-weekly.xlsx'
        outputFilePath = os.path.join(directory, out_filename)

        return outputFilePath

    def successful(self):
        mb.showinfo('成功', 'weekly交期已生成并存于%s' % self.gen_output_file_path())

    def start_convert(self):

        input = inputFilePath
        output = self.gen_output_file_path()

        print(input, output)

        if input is None:
            mb.showerror('模板错误', '请选择weekly模板')

        if self.cutOffDay is None:
            mb.showerror('日期错误', '请选择Cut Off Day')

        else:
            handler = Excel_handler(input, self.cutOffDay)
            handler.write_to_excel(output)
            # handler.test_save()
            self.successful()


if __name__ == '__main__':

    root = Tk()
    root.title('Daily2Weekly Tool')

    root.geometry('500x400')
    root.resizable(0, 0)

    input = Input(root)
    output = Output(root)

    root.rowconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)
    root.rowconfigure(2, weight=1)
    root.rowconfigure(3, weight=1)
    root.rowconfigure(4, weight=1)
    root.rowconfigure(5, weight=1)
    root.rowconfigure(6, weight=1)
    root.rowconfigure(7, weight=1)
    root.columnconfigure(0, weight=1)
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)
    root.columnconfigure(3, weight=1)

    root.mainloop()
