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
    def __init__(self, window, yr, month, day, 
                font='Constantia', f_size=12,
                ):
        self.yr = yr
        self.month = month
        self.day = day
        self.f = font
        self.f_size = f_size   
        self.font = (font, f_size)
        self.window = window
    
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

def program(start_cal, start_hr_entry, start_min_entry, start_sec_entry,
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
    #try:
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

    doc.createDoc(add_pic, variants_dict, payment_dict, pay_dict, t_title=t_title, product=product, images=images, pro_categ=pro_categ)
    doc.createPath(report_name)
    #except KeyError as e:
        #messagebox.showinfo("No results for the selected period", e)
    #except Exception as e:
        #messagebox.showinfo("Error", e)
    return

def main():
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root = ThemedTk(theme="breeze")

    try:
        root.iconbitmap('logo.ico')
    except:
        pass

    root.title('Sales Report')

    w = 600
    h = 150

    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()

    x = ((ws/2) - (w/2))
    y = ((hs/2) - (h/2))

    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
    root.columnconfigure(1, weight=1)

    start_yr, start_month, start_day, start_hour, start_min, start_sec, end_yr, end_month, end_day, end_hour, end_min, end_sec = rp.get_time()

    start_v = View(root, start_yr, start_month, start_day)
    end_v = View(root, end_yr, end_month, end_day)

    start_cal = start_v.date_range('From: ', 0, 0, 0, 1, 7)
    end_cal = end_v.date_range('to: ', 1, 0, 1, 1, 3)

    start_hr_entry = start_v.time_range(start_hour, 0, 23, 4, 0)
    start_min_entry = start_v.time_range(start_min, 0, 60, 5, 0)
    start_sec_entry = start_v.time_range(start_sec, 0, 60, 6, 0)

    end_hr_entry = end_v.time_range(end_hour, 0, 23, 4, 1)
    end_min_entry = end_v.time_range(end_min, 0, 60, 5, 1)
    end_sec_entry = end_v.time_range(end_sec, 0, 60, 6, 1)

    ttk.Button(root, text="Print", command=lambda: program(start_cal, start_hr_entry, start_min_entry, start_sec_entry, end_cal, end_hr_entry, end_min_entry, end_sec_entry), style='print.TButton').grid(column=5, row=6, padx=10, columnspan=3)

    root.mainloop()
    return root, start_cal, start_hr_entry, start_min_entry, start_sec_entry, end_cal, end_hr_entry, end_min_entry, end_sec_entry

if  __name__ == '__main__':
    main()





