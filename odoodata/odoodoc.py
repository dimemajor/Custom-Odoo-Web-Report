import datetime as dt
import os

import glob
import docx
from docx2pdf import convert
from docx.enum.text import WD_BREAK
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

from constants import *
import odoodata.odoodata as od


class Doc():
    def __init__(self, start_date=None, end_date=None, 
                start_time=None, end_time=None, h_font='Alef'
            ):

        cwd = os.getcwd()
        if '/' in cwd:
            cwd = cwd.split('/')
        elif '\\' in cwd:
            cwd = cwd.split('\\')
        if ':' in cwd[0]:
            self.path = f'{cwd[0]}/Users/{cwd[2]}/Documents/'
            self.path_to_catalog = f'{cwd[0]}/Users/{cwd[2]}/Documents/MARMOR CATALOG/'
        else:
            self.path = f'{cwd[0]}/Users/{cwd[2]}/Documents/'
            self.path_to_catalog = f'Users/{cwd[0]}/Documents/MARMOR CATALOG/'

        self.comp = od.getData(PASSWORD, EMAIL, DOMAIN, start_date, end_date, start_time, end_time)
        self.h_font = h_font
        self.doc = docx.Document()

        if start_date is None and end_date is None:
            self.start = dt.date.today()
            self.date = f'{self.start:%d %b %y}'
        elif start_date is None and end_date != None:
            date = end_date.split('-')
            date = [int(value) for value in date]
            self.start = dt.date(date[0],date[1],date[2])
            self.date = f'{self.start:%d %b %y}'
        elif start_date != None and end_date is None:
            date = start_date.split('-')
            date = [int(value) for value in date]
            self.start = dt.date(date[0],date[1],date[2])
            self.date = f'{date:%d %b %y}'
        elif start_date != None and end_date != None:
            start = start_date.split('-')
            start = [int(value) for value in start]
            start_time = start_time.split(':')
            start_time = [int(value) for value in start_time]
            self.start = dt.datetime(start[0],start[1],start[2],start_time[0]+1,start_time[1],start_time[0])

            end = end_date.split('-')
            end = [int(value) for value in end]
            end = [int(value) for value in end]
            end_time = end_time.split(':')
            end_time = [int(value) for value in end_time]
            self.end = dt.datetime(end[0],end[1],end[2],end_time[0]+1,end_time[1],end_time[2])
            self.date = f'{self.start:%d %b %y_%I-%M %p} - {self.end:%d %b %y_%I-%M %p}'

    def formatVariants(self):
        variants_dict = []
        temp_dict = {}
        product = []
        images = []
        pro_categ = []
        cost = []
        headings = ['Product', 'Qty', 'Unit Price', 'Total Price']

        response = self.comp.getVariantsSales()
        for record in response.json()['result']:
            for field, value in record.items():
                if field == 'product_id':
                    if ' ' in value[1]:
                        pro_split = value[1].split(' ')
                        if '[' not in pro_split[0]:
                            product.append(pro_split[0])
                            try:
                                images.append(pro_split[0]+' '+pro_split[1].title())
                                temp_dict[headings[0]] = pro_split[0]+' '+pro_split[1]
                            except:
                                images.append(pro_split[0])
                                temp_dict[headings[0]] = pro_split[0]
                        else:
                            product.append(pro_split[1])
                            try:
                                images.append(pro_split[1]+' '+pro_split[2].title())
                                temp_dict[headings[0]] = pro_split[1]+' '+pro_split[2]
                            except:
                                images.append(pro_split[1])
                                temp_dict[headings[0]] = pro_split[1]
                    else:
                        product.append(value[1])
                        images.append(value[1])
                        temp_dict[headings[0]] = value[1]
                elif field == 'product_qty':
                    temp_dict[headings[1]] = f'{value:,.2f}'
                elif field == 'average_price':
                    temp_dict[headings[2]] = f'{value:,.2f}'
                elif field == 'price_total':
                    cost.append(value)
                    temp_dict[headings[3]] = f'{value:,.2f}'
                elif field == 'product_categ_id':
                    pro_categ.append(value[1])
                    break
                else:
                    continue
            
            temp_dict = {key: temp_dict[key] for key in headings}
            variants_dict.append(temp_dict.copy())
        return variants_dict, product, images, pro_categ, cost

    def formatPayment(self, method):
        payment_dict = []
        cash_amount = []
        trfs_amount = []
        pos_amount = []
        cus_amount = []
        ids = []
        dep_increase = []
        dep_decrease = []

        response = self.comp.getPayments()
        for i in response.json()['result']['records']:
            if i['payment_method_id'][1] == 'Cash':
                cash_amount.append(i['amount'])
            elif i['payment_method_id'][1] == 'Transfer':
                trfs_amount.append(i['amount'])
            elif i['payment_method_id'][1] == 'POS':
                pos_amount.append(i['amount'])
            elif i['payment_method_id'][1] == 'Customer Account':
                ids.append(i['pos_order_id'][0])
                cus_amount.append(i['amount'])
        t_cash = sum(cash_amount)
        t_trfs = sum(trfs_amount)
        t_pos = sum(pos_amount)
        t_cus = sum(cus_amount)
        totals = t_cash, t_trfs, t_pos, t_cus

        response = self.comp.getOrders()
        j=0
        for i in response.json()['result']['records']:
            temp_dict = {}
            temp_dict['Date'] = i['date_order']
            try:
                temp_dict['Customer'] = i['partner_id'][1]
            except:
                temp_dict['Customer'] = 'General'
            try:
                x = ids.index(i['id'])
                pay = float(cus_amount[x])
            except:
                temp_dict['Amount'] = i['amount_total']

            if method == 'default':
                if temp_dict['Customer'] != 'General': #because of the type of data, if temp_dict['amount']==None, this line wont run. temp_dict['amount'] is already set
                    temp_dict['Amount'] = pay
                    payment_dict.append(temp_dict)

            elif method == 'custom':
                if j==0:
                    bal = 0
                if temp_dict['Customer'] != 'General':
                    if pay >= 0:
                        temp_dict['in'] = pay
                        dep_increase.append(pay)
                        temp_dict['out'] = ''
                        #temp_dict['balance'] = bal + pay  
                    else:
                        temp_dict['in'] = ''
                        temp_dict['out'] = abs(pay)
                        dep_decrease.append(abs(pay))
                        #temp_dict['balance'] = bal - abs(pay)
                    #bal = temp_dict['balance']
                    payment_dict.append(temp_dict)
                    j+=1
            else:
                raise ValueError('try "default" or "custom"') 
        return payment_dict, totals, dep_increase, dep_decrease

    def docHeader():
        def decorator(func):
            def wrapper(self, add_pic, *args, **kwargs):
                print('creating document...')
                section = self.doc.sections[0]
                header = section.header
                header_para = header.paragraphs[0]
                header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = header_para.add_run()
                font = run.font
                font.name = self.h_font
                font.size = Pt(15)
                run.text = COMPANY
                para = header.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.text = self.date
                func(self, add_pic, *args, **kwargs)
            return wrapper
        return decorator

    @docHeader()
    def createDoc(self, add_pic, *args, t_title=[], product=[], images=[],
        pro_categ=[], t_style='Table Grid', f_name='Alef',
        ):

        def get_image_hyperlink(img_path):
            hyp_name = 'file:///' + img_path
            return hyp_name

        def add_hyperlink(paragraph, text, url):
            part = paragraph.part
            r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

            hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
            hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

            new_run = docx.oxml.shared.OxmlElement('w:r')
            rPr = docx.oxml.shared.OxmlElement('w:rPr')

            new_run.append(rPr)
            new_run.text = text
            hyperlink.append(new_run)

            r = paragraph.add_run ()
            r._r.append (hyperlink)

            r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
            r.font.underline = True
            r.font.size = Pt(7)
            return hyperlink

        def find_image(img, img_v, category, path):
            image = glob.glob(f'{path}{category}/{img_v}*')
            if len(image)==0:
                image = glob.glob(f'{path}{category}/{img}*')
            return image[0]

        if len(args) != len(t_title):
            raise ValueError('Length of table and table title must match')
        if len(args) != len(add_pic):
            raise ValueError('Length of bool list must match table length')

        for n, table in enumerate(args):
            self.doc.add_heading(t_title[n], 1)
            if add_pic[n] == True:
                if len(product) > len(table):
                    raise ValueError('Length of table must be >= product length')
                if len(images) > len(table):
                    raise ValueError('Length of table must be >= images length')
                if len(pro_categ) > len(table):
                    raise ValueError('Length of table must be >= product category length')
                t = self.doc.add_table(len(table)+1, len(table[0].keys())+1)
                t.cell(0,len(table[0].keys())).text = 'Image'
            else:
                t = self.doc.add_table(len(table)+1, len(table[0].keys()))
            
            t.style = t_style
            font = t.style.font
            font.name = f_name

            j=0
            for key in table[0].keys():
                t.cell(0,j).text = key
                j+=1
            i=1
            for record in table:
                j=0
                for value in record.values():
                    if isinstance(value, int) or isinstance(value, float):
                        value = f'{value:,.2f}'
                    t.cell(i,j).text = value
                    j+=1
                i+=1
            i=1
            for record in range(len(images)):
                if add_pic[n] == True:
                    try:
                        cell = t.cell(i,len(table[0].keys()))
                        paragraph = cell.paragraphs[0]
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = paragraph.add_run()
                        img_path = find_image(product[i-1], images[i-1], pro_categ[i-1], self.path_to_catalog)
                        run.add_picture(img_path, width = Inches(.7), height = Inches(.8))
                        run.add_break(WD_BREAK.LINE)
                        hyp = get_image_hyperlink(img_path)
                        add_hyperlink(paragraph, 'open', hyp)
                        
                        paragraph = cell.add_paragraph()
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        file_path, file_ext = os.path.splitext(img_path)
                        mac_image = f'file:///Users/morufat/Documents/MARMOR CATALOG/{pro_categ[i-1]}/{product[i-1]}{file_ext}'
                        add_hyperlink(paragraph, ' photo', mac_image)
                    except:
                        t.cell(i,len(table[0].keys())).text = 'no image'
                i+=1
        
    def createPath(self, report_name):
        def create_file(new_month_dir):
            #keep trying to delete doc and set status==1(a good thing) to proceed
            doc_in_use = True
            i=1
            while doc_in_use:
                if os.path.exists(f"{new_month_dir}/{report_name} ({i}).docx") or os.path.exists(f"{new_month_dir}/{report_name} ({i}).pdf"):
                    try:
                        os.remove(f"{new_month_dir}/{report_name} ({i}).docx")
                        status=1
                    except FileNotFoundError:
                        status=1
                        pass
                    except:
                        status=0
                        i+=1
                    if status==1:
                        try:
                            os.remove(f"{new_month_dir}/{report_name} ({i}).pdf")
                        except:
                            i+=1
                    
                else:
                    doc_in_use = False
            return i

        pro_dir = rf'{self.path}/{FOLDER_NAME}'
        try:
            new_month_dir = rf'{self.path}/{FOLDER_NAME}/{self.start:%b %y}'
        except:
            new_month_dir = rf'{self.path}/{FOLDER_NAME}/{self.date:%b %y}'
        #print(new_month_dir)
        if not os.path.exists(pro_dir):
            os.mkdir(pro_dir)
        if not os.path.exists(new_month_dir):
            os.mkdir(new_month_dir)
        i=0
        if os.path.exists(f"{new_month_dir}/{report_name}.docx") or os.path.exists(f"{new_month_dir}/{report_name}.pdf"):
            try:
                os.remove(f"{new_month_dir}/{report_name}.docx")
                status=1 # 1 means its a good thing and is expected
            except FileNotFoundError:
                status=1
                pass
            except:
                status=0 # 0 means not what was expected. Maybe the doc to be deleted is open or cannot delete for some reason
                i=create_file(new_month_dir) # so enter a loop and keep trying with a different name until status==1
            if status==1:
                try:
                    os.remove(f"{new_month_dir}/{report_name}.pdf")
                except:
                    i=create_file(new_month_dir)
        if i == 0:
            report_path = f'{new_month_dir}/{report_name}'
        else:
            report_path = f'{new_month_dir}/{report_name} ({i})'

        print('saving and converting doc..')
        self.doc.save(f'{report_path}.docx')
        convert(f'{report_path}.docx')
        try:
            os.remove(f'{report_path}.docx')
            os.startfile(f'{report_path}.pdf')
        except:
            pass


def get_time():

    today = dt.datetime.today()
    yesterday = dt.datetime.today()-dt.timedelta(1)
    try:
        comp = od.Session(PASSWORD, EMAIL, DOMAIN)
        start_yr, start_month, start_day, start_hour, start_min, start_sec = format_pos(comp)
    except:
        start_day = int(f'{yesterday:%d}')
        start_month =  int(f'{yesterday:%m}')
        start_yr = int(f'{yesterday:%y}')
        start_hour = int(f'{yesterday:%H}')
        start_min = int(f'{yesterday:%M}')
        start_sec = int(f'{yesterday:%S}')
        
    end_day = int(f'{today:%d}')
    end_month =  int(f'{today:%m}')
    end_yr = int(f'{today:%y}')
    end_hour = int(f'{today:%H}')
    end_min = int(f'{today:%M}')
    end_sec = int(f'{today:%S}')
    
    return start_yr, start_month, start_day, start_hour, start_min, start_sec, end_yr, end_month, end_day, end_hour, end_min, end_sec

def format_pos(comp):
    response = comp.getPosData()
    time = response.json()['result']['records'][0]['start_at']
    time_split = time.split(' ')

    td = time_split[0].split('-')
    start_yr = int(td[0])
    start_month = int(td[1])
    start_day = int(td[2])

    tt = time_split[1].split(':')
    hour = int(tt[0])+1
    min = int(tt[1])
    sec = int(tt[2])
    return start_yr, start_month, start_day, hour, min, sec



