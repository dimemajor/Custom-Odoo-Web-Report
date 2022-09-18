import ctypes
import babel.numbers
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from ttkthemes import ThemedTk
from tkinter import messagebox
import datetime as dt

import odoodata.odoodoc as rp


class View():
    def __init__(self, window, yr=None, month=None, day=None, 
                font='Constantia', f_size=12,
                ):
        self.yr = yr
        self.month = month
        self.day = day
        self.f = font
        self.f_size = f_size   
        self.font = (font, f_size)
        self.window = window
        self.style = ttk.Style().configure('small.TButton', font=self.font)

    
    def date_range(self, text, label_row, label_column, entry_row, entry_column, pady):
        cal_label = tk.Label(self.window, text=text, font=self.font)
        cal_label.grid(column=label_column, row=label_row, pady=pady, sticky='ew')
        cal = DateEntry(self.window, year=self.yr, font=self.font, month=self.month, day=self.day, width=12, background='grey',foreground='white', borderwidth=1)
        cal.grid(row=entry_row, column=entry_column,sticky='ew',columnspan=3, pady=pady)
        return cal

    def time_range(self, time, from_, to, col, row):
        hr_var= tk.StringVar(self.window)
        hr_var.set(time)
        hr_entry = tk.Spinbox(self.window,textvariable=hr_var,from_=from_,to=to,width=3, font=self.font)
        hr_entry.grid(column=col, row=row,pady=7, sticky='e',columnspan=1)
        return hr_entry

    def add_button(self, func, row, col, colspan=1, padx=5, pady=5, text='Print'):
        ttk.Button(self.window, text=text, command=func, style=self.style).grid(row=row, column=col, columnspan=colspan, padx=padx, pady=pady)
        return
        

def sales_report(start_cal, start_hr_entry, start_min_entry, start_sec_entry,
                end_cal, end_hr_entry, end_min_entry, end_sec_entry):
    start_date = str(start_cal.get_date())
    end_date = str(end_cal.get_date())

    start_hr = int(start_hr_entry.get())
    start_min = int(start_min_entry.get())
    start_sec = int(start_sec_entry.get())

    end_hr = int(end_hr_entry.get())
    end_min = int(end_min_entry.get())
    end_sec = int(end_sec_entry.get())

    start_time = f'{str(start_hr-1)}:{str(start_min)}:{str(start_sec)}'
    end_time = f'{str(end_hr-1)}:{str(end_min)}:{str(end_sec)}'

    t1 = start_date.split('-')
    t2 = end_date.split('-')
    
    temp_start = [int(value) for value in t1]
    temp_end = [int(value) for value in t2]
    
    start = dt.datetime(temp_start[0],temp_start[1],temp_start[2],start_hr,start_min)
    end = dt.datetime(temp_end[0],temp_end[1],temp_end[2],end_hr,end_min)
    
    report_name = f'{start:%d %b %y} - {end:%d %b %y}'
    try:
        doc = rp.Doc(start_date=start_date, end_date=end_date, start_time=start_time, end_time=end_time)
        variants_dict, product, images, pro_categ, cost = doc.formatVariants()
        payment_dict, totals, dep_increase, dep_decrease = doc.formatPayment('custom')

        pay_dict = [
            {
                'Name':'Cash',
                'Total':totals[0],
            },
            {
                'Name':'Transfer',
                'Total':totals[1],
            },
            {
                'Name':'POS',
                'Total':totals[2],
            },
            {
                'Name':'Total',
                'Total':totals[0] + totals[1] + totals[2]
            }
        ]
        total_dep_dict = {
            'Date' : 'Total',
            'Customer' : '',
            'in': sum(dep_increase),
            'out': sum(dep_decrease),
        }

        payment_dict.append(total_dep_dict)
        t_title = ['Products', 'Customer Deposits', 'Payments']
        add_pic = [True, False, False]
        dicts = [variants_dict, payment_dict, pay_dict]
        doc.createDoc(add_pic, dicts, t_title=t_title, product=product, images=images, pro_categ=pro_categ)
        doc.createPath(report_name)
    except KeyError as e:
        messagebox.showinfo("No results for the selected period", e)
    except Exception as e:
        messagebox.showinfo("Error", e)
    return

def select_all(lb, list):
    if len(lb.curselection()) == len(list):
        lb.selection_clear(0, tk.END)
    elif len(lb.curselection()) > 1:
        lb.select_set(0, tk.END)
    else:
        lb.select_set(0, tk.END)

def get_quants(lb):
    titles=[]
    categs=[]
    add_pic=[]
    dicts=[]
    for i in lb.curselection():
        categs.append(lb.get(i))
    invt_list = rp.Doc()
    shop_list, title, pic = invt_list.getUpdatedQuant(location='shop', categs=categs)
    dicts.extend(shop_list)
    titles.extend(title)
    add_pic.extend(pic)
    wh_list, title, pic = invt_list.getUpdatedQuant(location='warehouse', categs=categs)
    dicts.extend(wh_list)
    titles.extend(title)
    add_pic.extend(pic)

    if shop_list is not None and wh_list is not None:
        invt_list.createDoc(add_pic, dicts, t_title=titles)
    elif wh_list is not None:
        invt_list.createDoc(add_pic, wh_list, t_title=titles)
    elif shop_list is not None:
        invt_list.createDoc(add_pic, shop_list, t_title=titles)
    else:
        messagebox.showinfo('Info', 'No results to display')
    
    invt_list.createPath(report_name = f'Inventory List')

def main():
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root = ThemedTk(theme="breeze")

    try:
        root.iconbitmap('logo.ico')
    except:
        pass

    root.title('Sales Report')
    w = 650
    h = 185

    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()

    x = ((ws/2) - (w/2))
    y = ((hs/2) - (h/2))

    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    notebook = ttk.Notebook(root, width=w, height=h)
    notebook.pack()

    repo_frame = ttk.Frame(notebook, width=w, height=h)
    invt_frame = ttk.Frame(notebook, width=w, height=h)
    repo_frame.pack(fill='both', expand=True)
    invt_frame.pack(fill='both', expand=True)
    repo_frame.columnconfigure(1, weight=1)
    invt_frame.columnconfigure(1, weight=1)

    notebook.add(repo_frame, text='Sales Report')
    notebook.add(invt_frame, text='Inventory')

    start_yr, start_month, start_day, start_hour, start_min, start_sec, end_yr, end_month, end_day, end_hour, end_min, end_sec = rp.get_time()

    start_v = View(repo_frame, start_yr, start_month, start_day)
    end_v = View(repo_frame, end_yr, end_month, end_day)

    start_cal = start_v.date_range('From: ', label_row=0, label_column=0, entry_row=0, entry_column=1, pady=7)
    end_cal = end_v.date_range('to: ', label_row=1, label_column=0, entry_row=1, entry_column=1, pady=3)

    start_hr_entry = start_v.time_range(start_hour, from_=0, to=23, col=4, row=0)
    start_min_entry = start_v.time_range(start_min, from_=0, to=60, col=5, row=0)
    start_sec_entry = start_v.time_range(start_sec, from_=0, to=60, col=6, row=0)

    end_hr_entry = end_v.time_range(end_hour, from_=0, to=23, col=4, row=1)
    end_min_entry = end_v.time_range(end_min, from_=0, to=60, col=5, row=1)
    end_sec_entry = end_v.time_range(end_sec, from_=0, to=60, col=6, row=1)

    #ttk.Button(repo_frame, text="Print", command=lambda: program(start_cal, start_hr_entry, start_min_entry, start_sec_entry, end_cal, end_hr_entry, end_min_entry, end_sec_entry), style='small.TButton').grid(column=5, row=6, padx=10, columnspan=3)
    
    sales_buttn = View(repo_frame, f_size=9)
    sales_buttn.add_button(func=lambda: sales_report(start_cal, start_hr_entry, start_min_entry, start_sec_entry, end_cal, end_hr_entry, end_min_entry, end_sec_entry), row=6, col=5, colspan=3, padx=10)

    items = ['NET LACE', 'SEQUENCE LACE', 'BRIDAL LACE']
    categs = tk.Variable(value=items)
    categ_listbox = tk.Listbox(invt_frame, listvariable=categs, height=4, width=80, selectmode=tk.MULTIPLE, font=('Constantia', 9))
    categ_listbox.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

    invt_buttn = View(invt_frame, f_size=9)
    invt_buttn.add_button(func=lambda: select_all(categ_listbox, items), text='All/None', row=7, col=0, padx=10)
    invt_buttn.add_button(func=lambda: get_quants(categ_listbox), text= 'Print', row=7, col=2, padx=10)

    root.mainloop()
    return root, start_cal, start_hr_entry, start_min_entry, start_sec_entry, end_cal, end_hr_entry, end_min_entry, end_sec_entry

if  __name__ == '__main__':
    main()





