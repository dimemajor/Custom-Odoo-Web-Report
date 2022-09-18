import datetime as dt
import os

import pandas as pd
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
from utils import *


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
            self.path_to_catalog = f'{cwd[0]}/Users/{cwd[2]}/Documents/{CATALOG_FOLDER}/'
        else:
            self.path = f'{cwd[0]}/Users/{cwd[2]}/Documents/'
            self.path_to_catalog = f'Users/{cwd[0]}/Documents/{CATALOG_FOLDER}/'

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
        ordered_list = ['Product', 'Qty', 'Unit Price', 'Total Price']

        response = self.comp.getVariantsSales()
        v_df = pd.DataFrame(response.json()['result'])
        to_keep = ['product_id', 'product_qty', 'average_price', 'price_total', 'product_categ_id']
        v_df = rm_headers(v_df, to_keep)
        v_df['product_categ_id'] = v_df['product_categ_id'].apply(replace)
        v_df['product_id'] = v_df['product_id'].apply(replace)
        v_df['product_id'] = v_df['product_id'].apply(rm_other_fields)
        v_df[['images', 'colour']] = v_df['product_id'].str.split(' ', expand=True)
        product = v_df['product_id'].to_list()
        pro_categ = v_df['product_categ_id'].to_list()
        images = v_df['images'].to_list()

        response = self.comp.getVariants()
        df = pd.DataFrame(response.json()['result']['records'])
        to_keep = ['categ_id', 'uom_id']
        df = rm_headers(df, to_keep)
        df['categ_id'] = df['categ_id'].apply(replace)
        df['uom_id'] = df['uom_id'].apply(replace)
        df = df.drop_duplicates(subset = "categ_id")
        df.rename(columns= {'categ_id': 'product_categ_id', 'uom_id': 'uom'}, inplace=True)
        df = pd.merge(v_df, df, on=['product_categ_id'], how='inner')
        df['product_qty'] = df['product_qty'].astype(float)
        df['product_qty'] = df.apply(pd_apply_conversion, axis=1)
        df.drop(['product_categ_id', 'uom', 'colour', 'images'], inplace=True, axis=1)
        df.rename(columns= {'product_qty': 'Qty', 'average_price': 'Unit Price', 'price_total': 'Total Price', 'product_id': 'Product'}, inplace=True)
        df = df[ordered_list]
        df.sort_values(by=['Product'], inplace=True)
        variants_dict = df.to_dict('records')
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
                        temp_dict['Out'] = pay
                        dep_increase.append(pay)
                        temp_dict['In'] = ''
                    else:
                        temp_dict['Out'] = ''
                        temp_dict['In'] = abs(pay)
                        dep_decrease.append(abs(pay))
                    payment_dict.append(temp_dict)
                    j+=1
            else:
                raise ValueError('try "default" or "custom"') 
        return payment_dict, totals, dep_increase, dep_decrease

    def variants_expanded(self):
        response = self.comp.getVariants()
        tem_df = pd.DataFrame(response.json()['result']['records'])
        to_keep = ['categ_id', 'name', 'uom_id', 'product_template_variant_value_ids']
        tem_df = rm_headers(df=tem_df, to_keep=to_keep)
        tem_df['categ_id'] = tem_df['categ_id'].apply(replace)
        tem_df['uom_id'] = tem_df['uom_id'].apply(replace)
        tem_df['product_template_variant_value_ids'] = tem_df['product_template_variant_value_ids'].apply((lambda x: replace(cell=x, index=0)))
        template_ids = tem_df['product_template_variant_value_ids'].to_list()
        template_ids = [x for x in template_ids if str(x) != 'nan']
        
        response = self.comp.attachVariants(template_ids)
        col_df = pd.DataFrame(response.json()['result'])
        col_df['display_name'] = col_df['display_name'].apply(lambda x: x.replace('COLOUR: ', ''))
        col_df.rename(columns={'id': 'product_template_variant_value_ids'}, inplace=True)
        df = pd.merge(tem_df, col_df, on=['product_template_variant_value_ids'], how='left')
        df['display_name'].fillna('', inplace=True)
        df['name'] = df['name'] + ' ' + '(' + df['display_name'] + ')'
        df['name'] = df['name'].apply(lambda x: str(x).replace(" ()", ""))
        df.rename(columns={'name': 'product', 'categ_id': 'Category', 'uom_id': 'uom'}, inplace=True)
        to_keep = ['product', 'Category', 'uom']
        df = rm_headers(df=df, to_keep=to_keep)

        return df

    def getUpdatedQuant(self, location, categs=None):
        def create_final_df(q_df):
            q_df.rename(columns= {'quantity': 'product_qty', 'Category_x': 'product_categ_id', 'product': 'Product', 'uom_x': 'uom'}, inplace=True)   
            q_df['Quantity'] = q_df.apply(pd_apply_conversion, axis=1)
            q_df.sort_values(by=['Product'], inplace = True)
            q_df.rename(columns= {'product_categ_id': 'Category'}, inplace=True)   
            to_keep = ['Product', 'Category', 'Quantity']
            q_df = rm_headers(df=q_df, to_keep=to_keep)
            return q_df

        titles=[]
        dicts=[]
        add_pic=[]
        var_df = self.variants_expanded()

        response = self.comp.getQuants()
        quant_df = pd.DataFrame(response.json()['result']['records'])
        quant_df['product_id'] = quant_df['product_id'].apply(replace)
        quant_df['product_categ_id'] = quant_df['product_categ_id'].apply(replace)
        quant_df['location_id'] = quant_df['location_id'].apply(replace)
        quant_df['product_uom_id'] = quant_df['product_uom_id'].apply(replace)
        to_keep = ['product_id', 'product_categ_id', 'location_id', 'product_uom_id', 'quantity']
        quant_df = rm_headers(df=quant_df, to_keep=to_keep)
        quant_df.rename(columns={'product_id': 'product', 'product_categ_id': 'Category', 'location_id': 'location', 'product_uom_id': 'uom'}, inplace=True)

        if categs != None:
            temp_variants_df = var_df.loc[(var_df['Category'].isin(categs))]
            temp_quant_df = quant_df.loc[(quant_df['Category'].isin(categs))]
        else:
            temp_quant_df = quant_df
            temp_variants_df = var_df

        if location == 'shop':
            temp_quant_df = temp_quant_df.loc[temp_quant_df['location']=='SP/Stock']
            temp_quant_df = pd.merge(temp_variants_df, temp_quant_df, on=['product'], how='left')
            temp_quant_df['quantity'] = temp_quant_df['quantity'].fillna(0)
            if temp_quant_df.shape[0] == 0:
                return None
            else:
                q_df = create_final_df(q_df=temp_quant_df)
        elif location == 'warehouse':
            temp_quant_df = temp_quant_df.loc[temp_quant_df['location']=='WH/Stock']
            temp_quant_df = pd.merge(temp_variants_df, temp_quant_df, on=['product'], how='inner')
            if temp_quant_df.shape[0] == 0:
                return None
            else:
                q_df = create_final_df(q_df=temp_quant_df)
        categ_set = set(q_df['Category'].to_list())
        for categ in categs:
            if categ in categ_set:
                temp_q_df = q_df.loc[q_df['Category']==categ]
                temp_q_df.drop(columns=['Category'], axis=1, inplace=True)
                dicts.append(temp_q_df.to_dict('records'))
                titles.append(f'{categ.title()}-{location.title()}')
                add_pic.append(False)
            
        return dicts, titles, add_pic

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
    def createDoc(self, add_pic, dicts, t_title=[], product=[], images=[],
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

        def find_image(img, img_v, category):
            images = glob.glob(f'{self.path_to_catalog}{category}/{img_v}*')
            if len(images)==0:
                images = glob.glob(f'{self.path_to_catalog}{category}/{img}*')
            if len(images)==0:
                images = glob.glob(f'{self.path_to_catalog}**/{img}*', recursive=True) #check all the folders for any matching file that bears the name of the product. Including all variants.
                if len(images) > 0:
                    img_path = None
                    for im in images:
                        file_path, file_ext = os.path.splitext(im)
                        file_name = file_path.split('\\')
                        try:
                            os.rename(im, f'{self.path_to_catalog}{category}/{file_name[len(file_name)-1]}{file_ext}') # move all the found files to their appropriate folers
                        except: # if the file cannot be moved for some reason, quickly check if this is the image we're looking for and set the return value before moving on.
                            if img_v in file_ext:
                                img_path = im
                    if not img_path: # if img_path was not set during the loop, check the appropriate folders again for the variants and the product respectively
                        images = glob.glob(f'{self.path_to_catalog}{category}/{img_v}*')
                        if len(images) == 0:
                            images = glob.glob(f'{self.path_to_catalog}{category}/{img}*')
                    else:
                        return img_path
            return images[0]

        if len(dicts) != len(t_title):
            raise ValueError('Length of table and table title must match')
        if len(dicts) != len(add_pic):
            raise ValueError('Length of bool list must match table length')

        for n, table in enumerate(dicts):
            self.doc.add_heading(t_title[n], 2)
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
            if add_pic[n] == True:
                for record in range(len(images)):
                    try:
                        cell = t.cell(i,len(table[0].keys()))
                        paragraph = cell.paragraphs[0]
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = paragraph.add_run()
                        img_path = find_image(images[i-1], product[i-1], pro_categ[i-1])
                        run.add_picture(img_path, width = Inches(.7), height = Inches(.8))
                        run.add_break(WD_BREAK.LINE)
                        hyp = get_image_hyperlink(img_path)
                        add_hyperlink(paragraph, 'open', hyp)
                        
                        paragraph = cell.add_paragraph()
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        file_path, file_ext = os.path.splitext(img_path)
                        mac_image = f'file:///Users/morufat/Documents/{CATALOG_FOLDER}/{pro_categ[i-1]}/{product[i-1]}{file_ext}'
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



