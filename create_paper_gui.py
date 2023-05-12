from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import *
from tkinter import ttk
from create_paper import create


class App(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)

        title_label = ttk.Label(self, text="防雷题库自动组卷工具 ver0.1", font=("fangsong", 20))
        file_label = ttk.Label(self, text="题库文件")
        self.file_var = StringVar()
        file_entry = ttk.Entry(self, state="readonly", textvariable=self.file_var)
        open_button = ttk.Button(self, text="打开...", command=self.add_file)
        copies_label = ttk.Label(self, text="出题份数")
        self.copies_var = IntVar(value=1)
        copies_spin = ttk.Spinbox(self, from_=1, to=10, textvariable=self.copies_var)
        self.shuffle_var = BooleanVar()
        shuffle_check = ttk.Checkbutton(
            self, text="打乱选项", variable=self.shuffle_var, onvalue=True
        )
        confirm_button = ttk.Button(self, text="生成", command=self.create)
        close_button = ttk.Button(self, text="退出", command=self.quit)

        self.grid(column=0, row=0, columnspan=4, rowspan=5, padx=3, pady=3)
        title_label.grid(column=1, row=0, columnspan=2, pady=3)
        file_label.grid(column=0, row=1, sticky=E, pady=3)
        file_entry.grid(column=1, row=1, columnspan=2, sticky=(E, W), padx=3)
        open_button.grid(column=3, row=1, sticky=E, padx=3)
        copies_label.grid(column=0, row=2, sticky=E)
        copies_spin.grid(column=1, row=2, sticky=W, padx=3)
        shuffle_check.grid(column=1, row=3, sticky=W)
        confirm_button.grid(column=1, row=4)
        close_button.grid(column=2, row=4)

    def add_file(self):
        self.file_var.set(askopenfilename(filetypes=[("Excel files","*.xlsx")]))

    def create(self):
        xlsx_path = self.file_var.get()
        number_of_copies = self.copies_var.get()
        is_mess_up = self.shuffle_var.get()
        if not xlsx_path:
            showerror(message="必须选择一个题库！")
        else:
            create(xlsx_path, number_of_copies, is_mess_up)


myapp = App()
myapp.master.title("防雷题库自动组卷工具")
myapp.master.resizable(FALSE, FALSE)
myapp.mainloop()
