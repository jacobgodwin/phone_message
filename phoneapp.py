import tkinter as tk
from tkinter import ttk
import win32com.client as win32

class PhoneApp:

    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('260x400')
        self.root.title('Phone Message')
        self.root.iconbitmap(default='images\gator.ico')
        self.var_cb1 = ''
        self.var_cb2 = ''
        self.var_cb3 = ''
        self.var_cb4 = ''
        self.var_cb5 = ''
        self.var_cb6 = ''
        self.var_cb7 = ''
        self.var_cb8 = ''

        self.build_gui()

    def StoreCheckbutton(self):
        if self.cb1.get() == 1:
            self.var_cb1 = '[Telephoned]'
        else:
            pass
        if self.cb2.get() == 1:
            self.var_cb2 = '[Please Call]'
        else:
            pass
        if self.cb3.get() == 1:
            self.var_cb3 = '[Came to see you]'
        else:
            pass
        if self.cb4.get() == 1:
            self.var_cb4 = '[Will call again]'
        else:
            pass
        if self.cb5.get() == 1:
            self.var_cb5 = '[Wants to see you]'
        else:
            pass
        if self.cb6.get() == 1:
            self.var_cb6 = '[High Priority]'
        else:
            pass
        if self.cb7.get() == 1:
            self.var_cb7 = '[Returned your call]'
        else:
            pass
        if self.cb8.get() == 1:
            self.var_cb8 = '[Special Attention]'
        else:
            pass


    def build_gui(self):
        ## Top frame
        frame_top = ttk.Frame(self.root)
        frame_top.pack()

        self.button_to = ttk.Button(frame_top, width=5, text='To:', command=lambda:self.ResolveTo())
        self.button_to.grid(row=0, column=0, pady=(10,0))

        self.entry_to = ttk.Entry(frame_top)
        self.entry_to.grid(row=0, column=1, pady=(10,0), ipadx=15)

        self.label_caller = ttk.Label(frame_top, text='Caller:')
        self.label_caller.grid(row=1, column=0, pady=(0,3))

        self.entry_caller = ttk.Entry(frame_top)
        self.entry_caller.grid(row=1, column=1, pady=(0,3), ipadx=15)

        self.label_company = ttk.Label(frame_top, text='Company:')
        self.label_company.grid(row=2, column=0, pady=(0,3))

        self.entry_company = ttk.Entry(frame_top)
        self.entry_company.grid(row=2, column=1, pady=(0,3), ipadx=15)

        self.label_phone = ttk.Label(frame_top, text='Phone:')
        self.label_phone.grid(row=3, column=0, pady=(0,3))

        self.entry_phone = ttk.Entry(frame_top)
        self.entry_phone.grid(row=3, column=1, pady=(0,3), ipadx=15)

        ## Middle Frame
        message_frame = ttk.LabelFrame(self.root,text="Message Type:")
        message_frame.pack(padx=5, pady=5)

        self.cb1 = tk.IntVar()
        self.check1 = ttk.Checkbutton(message_frame, variable=self.cb1, onvalue=1, offvalue=0, text='Telephoned')
        self.check1.grid(row=0, column=0, sticky='w')

        self.cb2 = tk.IntVar()
        self.check2 = ttk.Checkbutton(message_frame, variable=self.cb2, text='Please Call')
        self.check2.grid(row=0, column=1, sticky='w')

        self.cb3 = tk.IntVar()
        self.check3 = ttk.Checkbutton(message_frame, variable=self.cb3, text='Came to see you')
        self.check3.grid(row=1, column=0, sticky='w')

        self.cb4 = tk.IntVar()
        self.check4 = ttk.Checkbutton(message_frame, variable=self.cb4, text='Will call again')
        self.check4.grid(row=1, column=1, sticky='w')

        self.cb5 = tk.IntVar()
        self.check5 = ttk.Checkbutton(message_frame, variable=self.cb5, text='Wants to see you')
        self.check5.grid(row=2, column=0, sticky='w')

        self.cb6 = tk.IntVar()
        self.check6 = ttk.Checkbutton(message_frame, variable=self.cb6, text="High Priority")
        self.check6.grid(row=2, column=1, sticky='w')

        self.cb7 = tk.IntVar()
        self.check7 = ttk.Checkbutton(message_frame, variable=self.cb7, text='Returned your call')
        self.check7.grid(row=3, column=0, sticky='w')

        self.cb8 = tk.IntVar()
        self.check8 = ttk.Checkbutton(message_frame, variable=self.cb8, text='Special Attention')
        self.check8.grid(row=3, column=1, sticky='w')

        ## Textbox frame
        text_frame = ttk.Frame(self.root)
        text_frame.pack()

        self.entry_message = tk.Text(text_frame, height=7, width=25, wrap='word')
        self.entry_message.grid(rowspan=6, columnspan=2, pady=(5,15), sticky='n,s,e,w')

        ## Bottom frame
        button_frame = ttk.Frame(self.root)
        button_frame.pack()

        self.button_send = ttk.Button(button_frame, text='Send', command=lambda:[self.StoreCheckbutton(),self.Emailer()])
        self.button_send.grid(row=0, column=0, pady=(0,15), ipadx=10, padx=(0,10))

        self.button_exit = ttk.Button(button_frame, text='Exit', command=self.close_window)
        self.button_exit.grid(row=0, column=1, pady=(0,15), ipadx=10, padx=(10,0))
    
    def ResolveTo(self):
        search_string = self.entry_to.get()
        outlook = win32.gencache.EnsureDispatch('Outlook.Application')
        recipient = outlook.Session.CreateRecipient(search_string)
        recipient.Resolve()
        ae = recipient.AddressEntry
        email_address = None

        if 'EX' == ae.Type:
            eu = ae.GetExchangeUser()
            email_address = eu.PrimarySmtpAddress

        if 'SMTP' == ae.Type:
            email_address = ae.Address

        self.entry_to.delete(0, 'end')
        self.entry_to.insert(0, email_address)
        return

    def Emailer(self):
        # import win32com.client as win32

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = self.entry_to.get()
        mail.Subject = "Phone Message from Caller"
        mail.Body = 'Caller:  ' + self.entry_caller.get() + '\nCompany:  ' + self.entry_company.get() + '\nPhone:  ' + self.entry_phone.get() + '\n\nMessage:  ' + self.var_cb1 + '   ' + self.var_cb2 + '   ' + self.var_cb3 + '   ' + self.var_cb4 + '   ' + self.var_cb5 + '   ' + self.var_cb6 + '   ' + self.var_cb7 + '   ' + self.var_cb8 + '\n\n' + self.entry_message.get('1.0', 'end')
        mail.Display(True)


    def close_window(self):
        self.root.destroy()

    def main(self):
        self.root.mainloop()

if __name__ == '__main__':
        app = PhoneApp()
        app.main()