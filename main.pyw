import tkinter as tk
from tkinter import ttk
import time
import os
import csv
from configparser import ConfigParser
from tkinter import messagebox as mbox
from tkinter import filedialog as fdial
from htmltopdf import make_qrcode, make_code128, create_file


class MainWin(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self)
        master.geometry(set_size(master, absolute=False, win_ratio=0.8))

        self.compcode = tk.StringVar()
        self.compname = tk.StringVar()
        self.compaddress = tk.StringVar()
        self.first = True
        
        
        self.master = master
        master.columnconfigure(1, weight=1)
        master.rowconfigure(0, weight=1)
        self.frameSide = ttk.Frame(master)
        self.frameSide.grid(row=0, column=0, sticky='nw', padx=(5, 0))
        self.frameMain = ttk.Frame(master)
        self.frameMain.grid(row=0, column=1, sticky='nswe', padx=5, pady=5)
        self.frameMain.columnconfigure(0, weight=1)
        self.frameMain.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(self.frameMain, selectmode="extended", height=30, columns=("date", "partnum","quantity", 
            "weight", "sentto", "invoice", "dock", "handling", "container"),
            displaycolumns="date partnum quantity weight sentto invoice dock handling container")
        self.tree.grid(row=0, column=0, sticky="nswe", padx=(5,5), pady=(5,5))
        self.vsb = ttk.Scrollbar(self.frameMain, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.frameMain, orient="horizontal", command=self.tree.xview)
        self.vsb.grid(row=0, column=1, sticky="nse")
        self.hsb.grid(row=1, column=0, sticky="sew")
        self.tree.config(yscrollcommand=lambda f, l: self.autoscroll(self.vsb, f, l))
        self.tree.config(xscrollcommand=lambda f, l:self.autoscroll(self.hsb, f, l))
        self.tree.heading("#0", text="Shipment")
        self.tree.heading("date", text="Date")
        self.tree.heading("partnum", text="Part number")
        self.tree.heading("quantity", text="Quantity")
        self.tree.heading("weight", text="Gross weight")
        self.tree.heading("sentto", text="Recipient")
        self.tree.heading("invoice", text="Invoice")
        self.tree.heading("dock", text="Dock number")
        self.tree.heading("handling", text="Handling code")
        self.tree.heading("container", text="Container type")

        self.tree.column("#0",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("date",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("partnum",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("quantity",minwidth=20, width=80, stretch=True, anchor="center")
        self.tree.column("weight",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("sentto",minwidth=20, width=150, stretch=True, anchor="center")
        self.tree.column("invoice",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("dock",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("handling",minwidth=20, width=100, stretch=True, anchor="center")
        self.tree.column("container",minwidth=20, width=80, stretch=True, anchor="center")

        self.tree.tag_configure('even', background='#d9dde2')

        self.btadd = ttk.Button(self.frameSide, text='Add', command=self.add)
        self.btadd.grid(row=0, column=0, pady=(10,0), padx=5)
        self.btcopy = ttk.Button(self.frameSide, text='Copy', command=self.copy)
        self.btcopy.grid(row=1, column=0, pady=(5,0), padx=5)
        self.btchange = ttk.Button(self.frameSide, text='Edit', command=self.edit)
        self.btchange.grid(row=2, column=0, padx=5, pady=(5,0))
        self.btdelete = ttk.Button(self.frameSide, text='Delete', command=self.delete)
        self.btdelete.grid(row=3, column=0, padx=5, pady=(5,0))
        self.btprint = ttk.Button(self.frameSide, text='Create PDF', command=self.createpdf)
        self.btprint.grid(row=4, column=0, padx=5, pady=(5,0))
        ttk.Separator(self.frameSide, orient='horizontal').grid(row=5, column=0, padx=(5,5), pady=(15,10), sticky='we')
        self.btsetting = ttk.Button(self.frameSide, text='Settings', command=self.settings)
        self.btsetting.grid(row=6, column=0, padx=5, pady=(5,0))
        self.btparts = ttk.Button(self.frameSide, text='Spare parts', command=self.parts)
        self.btparts.grid(row=7, column=0, padx=5, pady=(5,0))
        self.btinfo = ttk.Button(self.frameSide, text='About', command=self.getinfo)
        self.btinfo.grid(row=8, column=0, padx=5, pady=(5,0))

        self.tree.bind('<Double-Button-1>', self.edit)
        self.tree.bind('<KeyPress-Delete>', self.delete)

        self.read_file()
        self.get_from_ini()

    def get_from_ini(self):
        self.partlist=[]
        self.config = ConfigParser()
        if not os.path.isfile('config.ini'):
            self.check_inifile()
        self.config.read('config.ini')
        sect = 'DefaultInfo'
        self.compcode.set(self.config.get(sect, 'CompanyCode'))
        self.compname.set(self.config.get(sect, 'CompanyName'))
        self.compaddress.set(self.config.get(sect, 'CompanyAddress'))
        self.master.title(self.compname.get())
        
    def getinfo(self):
        self.getinfoaboutdev = CopyrightWin(self.master)
    
    def add(self):
        maxid = self.last_id()
        self.Win = EntryWin(self.master, self.tree, values=None, last=maxid)
        self.write_file()

    def last_id(self):
        if self.tree.get_children():
            idmax = max([int(i) for i in self.tree.get_children()])
            return idmax

    def edit(self, event=None):
        if event:
            w = event.widget
            item = w.focus()
        else:
            item = self.tree.focus()
        if item:
            maxid = self.last_id()
            values = self.tree.set(item)
            self.Win = EntryWin(self.master, self.tree, values, item, last=maxid, edit=True)
            self.write_file()

    def parts(self):
        win = PartWin(self.master)

    def copy(self):
        item = self.tree.focus()
        if item:
            maxid = self.last_id()
            values = self.tree.set(item)
            self.Win = EntryWin(self.master, self.tree, values, item, last=maxid)
            self.write_file()

    def delete(self, event=None):
        item = self.tree.selection()
        if item:
            text = "Do you want to delete '%s' ?" %(str(item))
            ans = mbox.askyesno(title="Delete", message=text)
            if ans:
                for i in item:
                    self.tree.delete(i)
                self.write_file()
                item1 = self.tree.get_children()[0]
                self.tree.selection_set(item1)

    def createpdf(self):
        items = self.tree.selection()
        if items:
            config = ConfigParser()
            config.read('config.ini')
            sect = 'DefaultInfo'
            path = config.get(sect, 'PDFDirectory')
            if path:
                valueslist = []
                for item in items:
                    values = self.tree.set(item)
                    values['compname'] = self.compname.get()
                    values['compcode'] = self.compcode.get()
                    values['compaddress'] = self.compaddress.get()
                    # Data which should be encoded to QR code
                    qrdata = '[)>?06?P{0}?Q{1}?1JUN{2}{3}?20L2{4}?7Q{5}GT?2S{6}??'.format(values['partnum'], values['quantity'], 
                        values['compcode'], item, values['handling'], values['weight'], values['invoice'])
                    qrname = make_qrcode(qrdata, item)
                    data = '1JUN' + values['compcode'] + item
                    code128name = make_code128(data, item)
                    values['qrname'] = qrname
                    values['code128name'] = code128name
                    code128text = 'UN ' + values['compcode'] + ' ' + item
                    values['code128text'] = code128text
                    valueslist.append(values)
                create_file(valueslist)
            else:
                text = 'Сначала выберите папку для сохранения PDF файл'
                mbox.showwarning(title="Ошибка", message=text)
                self.settings(setpath=True)
        else:
            text = 'Please, first select items using \nbutton CTRL + Left mouse button!'
            mbox.showwarning(title="Warning", message=text)
            


    def settings(self, setpath=False):
        setwin = SettingWin(self.master, setpath)

    def zebra(self):
        childs = self.tree.get_children()
        if childs:
            n=0
            for child in childs:
                n += 1
                if (n%2==0):
                    tag='even'
                else:
                    tag='odd'
                self.tree.item(child, tags=(tag,))

    def write_file(self):
        if self.tree.get_children():
            with open('data.csv', 'w', encoding='utf-8') as file:
                fieldnames = ["id", "date", "partnum","quantity", "weight", "sentto", "invoice", "dock", "handling", "container"]
                writer = csv.DictWriter(file, fieldnames=fieldnames, delimiter=";")
                writer.writeheader()
                for item in self.tree.get_children():
                    mydata = self.tree.set(item)
                    mydata["id"] = item
                    writer.writerow(mydata)
            self.zebra()

    def read_file(self):
        if os.path.isfile('data.csv'):
            with open('data.csv', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile, fieldnames = None, delimiter=";")
                for row in reader:
                    self.tree.insert("", "end", row["id"], text=row["id"], 
                        values=[row["date"], row["partnum"], row["quantity"], row["weight"], row["sentto"], 
                        row["invoice"], row["dock"], row["handling"], row["container"]])
            self.zebra()

    def get_info(self):
        if os.path.isfile('config.ini'):
            self.info={}
            self.config = ConfigParser()
            self.config.read('config.ini')
            sect = 'DefaultInfo'
            for item in self.config.items(sect):
                self.info[item[0]] = item[1]


    def autoscroll(self, sbar, first, last):
        """Hide and show scrollbar as needed."""
        first, last = float(first), float(last)
        if first <= 0 and last >= 1:
            sbar.grid_remove()
        else:
            sbar.grid()
            sbar.set(first, last)

    def check_inifile(self):
        text = '''[DefaultInfo]
ContainerType = 
CompanyCode = 
CompanyName = 
CompanyAddress = 
Plant/Dock = 
MaterialHandlingcode =
PDFDirectory = 

[Parts]'''
        file = open('config.ini', 'w')
        file.write(text)
        file.close()

    
class EntryWin(tk.Frame):
    
    def __init__(self, master, tree, values=None, item=None, last=None, edit=False):
        tk.Frame.__init__(self)
        self.top = tk.Toplevel()
        self.top.resizable(False, False)
        self.top.geometry(set_size(self.top, 280, 335))
        if edit:
            self.top.title("Edit")
        else:
            self.top.title("Add")
        self.frame = ttk.Frame(self.top)
        self.frame.pack()
        self.amount = tk.StringVar()

        self.tree = tree
        self.values = values
        self.item = item
        self.edit = edit
        self.last = last
        self.auto = True

        ttk.Label(self.frame, text='#', font=("Arial", 10)).grid(row=0, column=0, sticky='w', pady=(5,0))
        self.id = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.id.grid(row=0, column=1, pady=(5,0))
        ttk.Label(self.frame, text='Date', font=("Arial", 10)).grid(row=0, column=2, sticky='w', pady=(5,0))
        self.date = ttk.Entry(self.frame, width=10, font=("Arial", 10))
        self.date.grid(row=0, column=3, pady=(5,0))
        ttk.Label(self.frame, text='Invoice no.', font=("Arial", 10)).grid(row=1, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.invoice = ttk.Entry(self.frame, width=18, font=("Arial", 10))
        self.invoice.grid(row=1, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Part number', font=("Arial", 10)).grid(row=2, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.partnum = ttk.Combobox(self.frame, font=("Arial", 10), width=16)
        self.partnum.grid(row=2, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Quantity', font=("Arial", 10)).grid(row=3, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.quant = ttk.Entry(self.frame, textvariable=self.amount, width=18, font=("Arial", 10))
        self.quant.grid(row=3, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Recipient', font=("Arial", 10)).grid(row=4, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.comboCont = ttk.Combobox(self.frame, width=18,
                                      values=("Company 1", "Company 2"))
        self.comboCont.grid(row=4, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Gross weight, kg', font=("Arial", 10)).grid(row=5, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.weight = ttk.Entry(self.frame, width=18, font=("Arial", 10))
        self.weight.grid(row=5, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Handling code', font=("Arial", 10)).grid(row=6, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.handling = ttk.Combobox(self.frame, font=("Arial", 10), width=16)
        self.handling.grid(row=6, column=2, columnspan=2, pady=(5,0))

        ttk.Label(self.frame, text='Dock address', font=("Arial", 10)).grid(row=7, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.dock = ttk.Entry(self.frame, width=18, font=("Arial", 10))
        self.dock.grid(row=7, column=2, columnspan=2, pady=(5,0))
        ttk.Label(self.frame, text='Container type', font=("Arial", 10)).grid(row=8, column=0, columnspan=2, sticky='w', pady=(5,0))
        self.container = ttk.Combobox(self.frame, font=("Arial", 10), width=16)
        self.container.grid(row=8, column=2, columnspan=2, pady=(5,0))

        self.calclabel = ttk.Label(self.frame, text='Result: ', font=("Arial", 10, 'bold'))
        self.calclabel.grid(row=9, column=0, columnspan=4, sticky='w', padx=(0, 0), pady=(5,8))

        self.btadd = ttk.Button(self.frame, text='OK', command=self.add)
        self.btadd.grid(row=10, column=1, pady=(10,0), padx=5)
        self.btcancel = ttk.Button(self.frame, text='Cancel', command=self.cancel)
        self.btcancel.grid(row=10, column=3, padx=5, pady=(10,0))
        self.initialfill()

        self.top.bind('<Escape>', self.cancel)
        self.amount.trace('w', lambda name, index, mode, key=self.amount: self.calc(key))
        


        self.top.grab_set()
        master.wait_window(self.top)

    def calc(self, key):
                if self.partnum.get() and key.get().isdigit() and self.auto and not self.edit:
                        quant = int(self.partsdic[self.partnum.get()][1])
                        weight = float(self.partsdic[self.partnum.get()][0])
                        packweight = float(self.partsdic[self.partnum.get()][2])
                        self.val = int(int(self.amount.get())/quant)
                        totalweight = int(quant * weight + packweight)
                        self.calclabel.config(text='Result: {0} boxes with {1} kg'.format(self.val, totalweight))
                        self.weight.delete(0, 'end')
                        self.weight.insert(0, str(totalweight))


        


    def initialfill(self):
                self.partsdic = {}
                self.packls = []
                self.handlingls = []
                self.config = ConfigParser()
                self.config.read('config.ini')
                sect = 'DefaultInfo'
                for item in self.config.items('Parts'):
                        opt = item[1].split(';')
                        self.partsdic[item[0]] = opt
                        
                for k, v in self.partsdic.items():
                        self.handlingls.append(v[3])
                for k, v in self.partsdic.items():
                        self.packls.append(v[4])
                self.partlist = sorted(list(self.partsdic.keys()))
                self.packlist = sorted(self.packls)
                self.handlist = sorted(self.handlingls)
                self.partnum.config(values=self.partlist)
                self.handling.config(values=self.packlist)
                self.container.config(values=self.handlist)
                

                if self.edit:
                        self.id.insert(0, self.item)
                        self.date.insert(0, self.values['date'])
                        self.quant.insert(0, self.values['quantity'])
                        self.partnum.set(self.values['partnum'])
                        self.weight.insert(0, self.values['weight'])
                        self.invoice.insert(0, self.values['invoice'])
                        self.comboCont.set(self.values['sentto'])
                        self.dock.insert(0, self.values['dock'])
                        self.handling.set(self.values['handling'])
                        self.container.set(self.values['container'])

                elif self.values and not self.edit:
                        self.val = 1
                        self.auto = False
                        idnum = '{:0>9}'.format(str(self.last+1))
                        self.id.insert(0, idnum)
                        self.date.insert(0, self.values['date'])
                        self.quant.insert(0, self.values['quantity'])
                        self.partnum.set(self.values['partnum'])
                        self.weight.insert(0, self.values['weight'])
                        self.invoice.insert(0, self.values['invoice'])
                        self.comboCont.set(self.values['sentto'])
                        self.dock.insert(0, self.values['dock'])
                        self.handling.set(self.values['handling'])
                        self.container.set(self.values['container'])

                else:
                        if self.last:
                                idnum = '{:0>9}'.format(str(self.last+1))
                        else:
                                idnum = '{:0>9}'.format(str(1))

                        self.id.insert(0, idnum)
                        self.dock.insert(0, self.config.get(sect, 'Plant/Dock'))
                        self.date.insert(0, time.strftime('%d.%m.%Y'))
                self.invoice.focus_set()

    def add(self):
        self.auto = False
        if not self.edit and not self.values:
            self.quant.delete(0, 'end')
            self.quant.insert(0, self.partsdic[self.partnum.get()][1])
        rawdata = self.getvalues()
        n=0
        for val in rawdata:
            if not val:
                n += 1
        if n==0:
            if self.edit:
                if self.id.get() != self.item:
                    self.tree.delete(self.item)
                    self.tree.insert('', 'end', self.id.get(), text=self.id.get(), values = rawdata)
                else:
                    self.tree.set(self.item, 'date', rawdata[0])
                    self.tree.set(self.item, 'partnum', rawdata[1])
                    self.tree.set(self.item, 'quantity', rawdata[2])
                    self.tree.set(self.item, 'weight', rawdata[3])
                    self.tree.set(self.item, 'sentto', rawdata[4])
                    self.tree.set(self.item, 'invoice', rawdata[5])
                    self.tree.set(self.item, 'dock', rawdata[6])
                    self.tree.set(self.item, 'handling', rawdata[7])
                    self.tree.set(self.item, 'container', rawdata[8])
            else:
                iid = int(self.id.get())
                for n in range(self.val):
                    idd = '{:0>9}'.format(str(iid + n))
                    self.tree.insert('', 'end', idd, text=idd, values = rawdata)
            self.top.destroy()
        else:
            text = 'Fill all fields first!'
            mbox.showwarning(title="Warning", message=text)

    def cancel(self, event=None):
        self.top.destroy()

    def getvalues(self):
        date = self.date.get()
        quanti = self.quant.get()
        partnum = self.partnum.get()
        weight = self.weight.get()
        invoice = self.invoice.get()
        sentto = self.comboCont.get()
        dock = self.dock.get()
        handling = self.handling.get()
        container = self.container.get()

        return date, partnum, quanti, weight, sentto, invoice, dock, handling, container

class SettingWin(tk.Frame):

    def __init__(self, master, setpath=False):
        tk.Frame.__init__(self)
        self.top = tk.Toplevel()
        self.top.resizable(False, False)
        self.top.geometry(set_size(self.top, 314, 185))
        self.top.title("Settings")

        self.container = tk.StringVar()
        self.compcode = tk.StringVar()
        self.compname = tk.StringVar()
        self.compaddress = tk.StringVar()
        self.dock = tk.StringVar()
        self.handling = tk.StringVar()
        self.pdfdir = tk.StringVar()
        self.edit = False
        self.setpath = setpath

        self.frameMain = ttk.Frame(self.top)
        self.frameMain.grid(row=0, column=0, sticky='nswe', padx=(8,0))
        self.bottom = ttk.Frame(self.top)
        self.bottom.grid(row=1, column=0, sticky='nswe', padx=(8,0))

        ttk.Label(self.frameMain, text='Company:', font=("Arial", 9)).grid(row=0, column=0, sticky='w', pady=(5,0))
        self.name = ttk.Entry(self.frameMain, width=28, font=("Arial", 10))
        self.name.grid(row=0, column=1, pady=(5,0))
        ttk.Label(self.frameMain, text='DUNS number:', font=("Arial", 9)).grid(row=1, column=0, sticky='w', pady=(5,0))
        self.code = ttk.Entry(self.frameMain, width=28, font=("Arial", 9))
        self.code.grid(row=1, column=1, pady=(5,0))
        ttk.Label(self.frameMain, text='Address:', font=("Arial", 9)).grid(row=2, column=0, sticky='w', pady=(5,0))
        self.address = ttk.Entry(self.frameMain, width=28, font=("Arial", 9))
        self.address.grid(row=2, column=1, pady=(5,0))
        ttk.Label(self.frameMain, text='Recipient code:', font=("Arial", 9)).grid(row=3, column=0, sticky='w', pady=(5,0))
        self.dockname = ttk.Entry(self.frameMain, width=28, font=("Arial", 9))
        self.dockname.grid(row=3, column=1, pady=(5,0))
        #ttk.Label(self.frameMain, text='Dock address:', font=("Arial", 9)).grid(row=5, column=0, sticky='w', pady=(5,0))
        

        self.framepdf = ttk.Frame(self.frameMain)
        self.framepdf.grid(row=5, column=0, columnspan=2)
        ttk.Label(self.framepdf, text='Dir for PDF:       ', font=("Arial", 9)).grid(row=0, column=0, sticky='w', pady=(5,0))
        self.pdfaddress = ttk.Entry(self.framepdf, width=23, font=("Arial", 9))
        self.pdfaddress.grid(row=0, column=1, padx=(3, 0), pady=(5,0))
        self.pdfbtn = ttk.Button(self.framepdf, text='...', width=3, command=self.folder)
        self.pdfbtn.grid(row=0, column=2, padx=(5, 0), pady=(5, 0))

        self.skpSave = ttk.Button(self.bottom, text='Save', command=self.save_list)
        self.skpSave.grid(row=0, column=0, sticky='nse', padx=(118, 5), pady=(10, 3))
        self.btcancel = ttk.Button(self.bottom, text='Cancel', command=self.cancel)
        self.btcancel.grid(row=0, column=1, sticky='nse', padx=(5, 5), pady=(10, 3))

        self.top.bind('<Escape>', self.cancel)

        self.get_from_ini()

        self.top.grab_set()
        master.wait_window(self.top)

    def get_from_ini(self):
        self.partlist=[]
        self.config = ConfigParser()
        self.config.read('config.ini')
        sect = 'DefaultInfo'
        self.compcode.set(self.config.get(sect, 'CompanyCode'))
        self.compname.set(self.config.get(sect, 'CompanyName'))
        self.compaddress.set(self.config.get(sect, 'CompanyAddress'))
        self.dock.set(self.config.get(sect, 'Plant/Dock'))
        self.pdfdir.set(self.config.get(sect, 'PDFDirectory'))
        self.update_list()
        if self.setpath:
            self.pdfaddress.focus_set()
        else:
            self.name.focus_set()

    def update_list(self):
        self.name.delete(0, 'end')
        self.code.delete(0, 'end')
        self.address.delete(0, 'end')
        self.dockname.delete(0, 'end')
        self.pdfaddress.delete(0, 'end')
        self.name.insert(0, self.compname.get())
        self.code.insert(0, self.compcode.get())
        self.address.insert(0, self.compaddress.get())
        self.dockname.insert(0, self.dock.get())
        self.pdfaddress.insert(0, self.pdfdir.get())

    def folder(self):
        dirpath = fdial.askdirectory(initialdir='D:/', mustexist=False,
                                     parent=self.master, title='Choose directory')
        if dirpath:
            self.pdfaddress.delete(0, 'end')
            self.pdfaddress.insert(0, dirpath)
    
    
    def change(self, event=None):
        self.compcode.set(self.code.get())
        self.compname.set(self.name.get())
        self.compaddress.set(self.address.get())
        self.dock.set(self.dockname.get())
        self.update_list()

    def save_list(self):
        sect = 'DefaultInfo'
        self.config.set(sect,'CompanyCode', self.code.get())
        self.config.set(sect, 'CompanyName', self.name.get())
        self.config.set(sect, 'CompanyAddress', self.address.get())
        self.config.set(sect, 'Plant/Dock', self.dockname.get())
        self.config.set(sect, 'PDFDirectory', self.pdfaddress.get())
        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)
        self.top.destroy()

    def cancel(self, event=None):
        self.top.destroy()

class PartWin(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self)
        self.top = tk.Toplevel()
        self.top.resizable(False, False)
        self.top.geometry(set_size(self.top, 560, 260))
        self.top.title("Spare parts")
        
        self.edit = False
        self.config = ConfigParser()

        self.frame = ttk.Frame(self.top)
        self.frame.grid(row=0, column=0, sticky='nswe', padx=(2,5), pady=5)

        self.bottom = ttk.Frame(self.top)
        self.bottom.grid(row=1, column=0, sticky='nswe', padx=(2,5), pady=5)

        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(self.frame, height=6, columns=("weight", "quantity", "packweight", "packtype", "dock"),
            displaycolumns="weight quantity packweight packtype dock")
        self.tree.grid(row=0, column=0, columnspan=8, sticky="nswe", padx=(5,5), pady=(5,5))
        
        self.tree.heading("#0", text="Partnum")
        self.tree.heading("quantity", text="Qn-ty in box")
        self.tree.heading("weight", text="Part weight, kg")
        self.tree.heading("packweight", text="Box weight")
        self.tree.heading("packtype", text="Container")
        self.tree.heading("dock", text="Dock")
        
        self.tree.column("#0",minwidth=20, width=80, stretch=True, anchor="center")

        self.tree.column("weight",minwidth=20, width=70, stretch=True, anchor="center")
        self.tree.column("quantity",minwidth=20, width=70, stretch=True, anchor="center")
        self.tree.column("packweight",minwidth=20, width=80, stretch=True, anchor="center")
        self.tree.column("packtype",minwidth=20, width=70, stretch=True, anchor="center")
        self.tree.column("dock",minwidth=20, width=70, stretch=True, anchor="center")
        

        ttk.Label(self.frame, text="Partnum", anchor='center').grid(row=1, column=0, sticky='nswe', pady=(5, 0))
        ttk.Label(self.frame, text="Weight, kg", anchor='center').grid(row=1, column=1, sticky='nswe', pady=(5, 0))
        ttk.Label(self.frame, text="q-ty in box", anchor='center').grid(row=1, column=2, sticky='nswe', pady=(5, 0))
        ttk.Label(self.frame, text="Box weight", anchor='center').grid(row=1, column=3, sticky='nswe', pady=(5, 0))
        ttk.Label(self.frame, text="Container", anchor='center').grid(row=1, column=4, sticky='nswe', pady=(5, 0))
        ttk.Label(self.frame, text="Dock", anchor='center').grid(row=1, column=5, sticky='nswe', pady=(5, 0))
        
        self.entry1 = ttk.Entry(self.frame, width=10, font=("Arial", 10))
        self.entry1.grid(row=2, column=0, sticky='nswe', padx=(5, 0), pady=(5, 0))
        self.entry2 = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.entry2.grid(row=2, column=1, sticky='nswe', padx=(5, 0), pady=(5, 0))
        self.entry3 = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.entry3.grid(row=2, column=2, sticky='nswe', padx=(5, 0), pady=(5, 0))
        self.entry4 = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.entry4.grid(row=2, column=3, sticky='nswe', padx=(5, 5), pady=(5, 0))
        self.entry5 = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.entry5.grid(row=2, column=4, sticky='nswe', padx=(5, 5), pady=(5, 0))
        self.entry6 = ttk.Entry(self.frame, width=9, font=("Arial", 10))
        self.entry6.grid(row=2, column=5, sticky='nswe', padx=(5, 5), pady=(5, 0))

        self.btnadd = ttk.Button(self.frame, text="+", width=4, command=self.validate)
        self.btnadd.grid(row=2, column=6, sticky='ns', padx=(0, 5), pady=(5, 0))
        self.btndel = ttk.Button(self.frame, text="-", width=4, command=self.delete)
        self.btndel.grid(row=2, column=7, sticky='ns', padx=(0, 5), pady=(5, 0))

        self.skpSave = ttk.Button(self.bottom, text='Save', command=self.save)
        self.skpSave.grid(row=0, column=0, sticky='nse', padx=(380, 5), pady=(5, 3))
        self.btcancel = ttk.Button(self.bottom, text='Cancel', command=self.cancel)
        self.btcancel.grid(row=0, column=1, sticky='nse', padx=(5, 5), pady=(5, 3))

        self.tree.bind('<Double-1>', self.edit_part)
        self.tree.bind('<KeyPress-Delete>', self.delete)
        self.top.bind('<Escape>', self.cancel)
        self.entry1.bind('<Return>', self.set_focus)
        self.entry4.bind('<Return>', self.validate)
        self.top.bind('<Escape>', self.cancel)

        self.get_from_ini()


        self.top.grab_set()
        master.wait_window(self.top)

    def get_from_ini(self):
        self.parts={}
        self.config.read('config.ini')
        sect = 'Parts'
        if self.config.items(sect):
            for item in self.config.items(sect):
                opt = item[1].split(';')
                self.parts[item[0]] = opt

        self.update_list()
        
    def set_focus(self, event):
        w = event.widget
        print(event)
    
    def update_list(self):
        self.tree.delete(*self.tree.get_children())
        templist = sorted(list(self.parts.keys()))
        for li in templist:
            self.tree.insert('', 'end', li, text=li, values=self.parts[li])
        self.tree.focus_set()

    def edit_part(self, event=None):
        item = self.tree.focus()
        if item:
            self.edit = True
            self.clear()
            self.entry1.insert(0, item)
            self.entry2.insert(0, self.parts[item][0])
            self.entry3.insert(0, self.parts[item][1])
            self.entry4.insert(0, self.parts[item][2])
            self.entry5.insert(0, self.parts[item][3])
            self.entry6.insert(0, self.parts[item][4])

    def change(self):
        if self.edit:
            item = self.tree.focus()
            del self.parts[item]
            self.edit = False
        templist = [self.entry2.get().replace(',', '.'), self.entry3.get(), self.entry4.get().replace(',', '.'), self.entry5.get(), self.entry6.get()]
        self.parts[self.entry1.get()] = templist
        self.clear()
        self.update_list()

    def clear(self):
        self.entry1.delete(0, 'end')
        self.entry2.delete(0, 'end')
        self.entry3.delete(0, 'end')
        self.entry4.delete(0, 'end')
        self.entry5.delete(0, 'end')
        self.entry6.delete(0, 'end')

    def validate(self, event=None):
        if len(self.entry1.get())>0 and len(self.entry2.get())>0 and len(self.entry3.get())>0 and len(self.entry4.get())>0 and len(self.entry5.get())>0 and len(self.entry6.get())>0: 
            try:
                fnum = float(self.entry2.get().replace(',', '.'))
                fnum = float(self.entry3.get())
                fnum = float(self.entry4.get().replace(',', '.'))
                self.change()
            except:
                mbox.showwarning(title="Warning", message="Wrong value! Check fields again")  
        else:
            mbox.showwarning(title="Warning", message="Fill empty fields first!")
            


    def delete(self, event=None):
        item = self.tree.focus()
        if item:
            del self.parts[item]
            self.update_list()

    def save(self):
        self.config['Parts']={}
        for k, v in self.parts.items():
            values = ';'.join(v)
            self.config.set('Parts', k, values)
        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)
        self.top.destroy()

    def cancel(self, event=None):
        self.top.destroy()


class CopyrightWin(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self)
        self.top = tk.Toplevel()
        self.top.resizable(False, False)
        self.top.geometry(set_size(self.top, 200, 100))
        self.top.title("About")

        ttk.Label(self.top, text='Bob Zimor © 2017').grid(row=0, column=1, padx=(10,10), pady=(15, 0), sticky='nswe')
        ttk.Label(self.top, text='Tel: +99897 272 19 10').grid(row=1, column=1, padx=(10,10), pady=(5, 0), sticky='nswe')
        ttk.Label(self.top, text='Email: bobzimor@gmail.com').grid(row=2, column=1, padx=(10,10), pady=(5, 0), sticky='nswe')

        self.top.grab_set()
        master.wait_window(self.top)



def set_size(win, w=0, h=0, absolute=True, win_ratio=None):
    winw = win.winfo_screenwidth()
    winh = win.winfo_screenheight()
    if not absolute:
        w = int(winw * win_ratio)
        h = int(winh * win_ratio)
        screen = "{0}x{1}+{2}+{3}".format(w, h, str(int(winw*0.1)), str(int(winh*0.05)))
    else:
        screen = "{0}x{1}+{2}+{3}".format(w, h, str(int((winw-w)/2)), str(int((winh-h)/2)))
    return screen


root = tk.Tk()
app = MainWin(root)
#root.iconbitmap(default='icon.ico')
root.mainloop()
