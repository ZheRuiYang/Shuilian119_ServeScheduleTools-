import tkinter as tk, shelve, re, os
import tkinter.messagebox as msgbox

class PIDInput(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.DatabasePath = os.path.abspath(__file__)
        self.pack()
        self.inputWidget()
        self.bind_all('<Return>', self.checkData)

    @property
    def DatabasePath(self):
        return self._DB
    @DatabasePath.setter
    def DatabasePath(self, val):
        val = os.path.split(val)[0]
        self._DB = os.path.join(val, 'PID')

    def inputWidget(self):
        self.idLabel = tk.Label(self, text='人員代號：')
        self.idEntry = tk.Entry(self)
        self.passLabel = tk.Label(self, text='登入密碼：')
        self.passEntry = tk.Entry(self, show='*')
        self.passLabel2 = tk.Label(self, text='確認登入密碼：')
        self.passEntry2 = tk.Entry(self, show='*')
        self.downLabel = tk.Label(self, text='下鍵點按次數：')
        self.downEntry = tk.Entry(self)
        self.confirm = tk.Button(self, text='確認', command=self.checkData)
        
        self.idLabel.grid(column=0, row=0, padx=25)
        self.idEntry.grid(column=1, row=0, padx=25)
        self.passLabel.grid(column=0, row=1)
        self.passEntry.grid(column=1, row=1)
        self.passLabel2.grid(column=0, row=2)
        self.passEntry2.grid(column=1, row=2)
        self.downLabel.grid(column=0, row=3)
        self.downEntry.grid(column=1, row=3)
        self.confirm.grid(column=0, columnspan=2)
        
        self.idEntry.focus_set()

    def checkData(self, event):
        self.pid = self.idEntry.get()
        self.pWord = self.passEntry.get()
        self.pWord2 = self.passEntry2.get()
        self.dTimes = int(self.downEntry.get())

        if self.pid and self.pWord and self.pWord2 and self.dTimes:
            with shelve.open(self.DatabasePath) as _id:
                _id[self.pid] = self.pWord
                _id[f'_{self.pid}'] = self.pWord2
                _id[f'{self.pid}DownTimes'] = self.dTimes
                raise SystemExit

    @property
    def pid(self):
        return self._id
    @pid.setter
    def pid(self, val):
        if not re.search(r'\w*', val):
            msgbox.showerror(title='人員代號錯誤',
                             message='人員代號只能是不含空白的中、英文或數字組合')
            self.idEntry.select_range(0, tk.END)
        else:
           self._id = val

    @property
    def pWord(self):
        return self._password
    @pWord.setter
    def pWord(self, val):
        if re.match(r'\w\d{9}', val):
            self._password = val
        else:
            msgbox.showerror(title='登入密碼設定錯誤',
                             message='請輸入真正能登入的密碼')
            self.passEntry.select_range(0, tk.END)

    @property
    def dTimes(self):
        return self._time
    @dTimes.setter
    def dTimes(self, val):
        if isinstance(val, int):
            self._time = val
        else:r
            msgbox.showerror(title='下鍵點按次數錯誤',
                             message='下鍵點按次數必須是個數字')
            self.downEntry.select_range(0, tk.END)

    @property
    def pWord2(self):
        return self._password2
    @pWord2.setter
    def pWord2(self, val):
        if val == self.passEntry.get():
            if val[0].isupper():
                val = val[0].lower() + val[1:]
            else:
                val = val[0].upper() + val[1:]
            self._password2 = val
        else:
            msgbox.showerror(title='確認密碼錯誤',
                             message='確認密碼與登入密碼不符')
            self.passEntry2.select_range(0, tk.END)

if __name__ == '__main__':
    app = PIDInput()
    app.master.title('新增人員')
    app.mainloop()
