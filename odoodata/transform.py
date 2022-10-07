import datetime as dt
import os

import pandas as pd

from utils import *
from constants import *
import odoodata.odoodata as od


def formatVariants(comp):
    variants_dict = []
    temp_dict = {}
    product = []
    images = []
    pro_categ = []
    cost = []
    ordered_list = ['Product', 'Qty', 'Unit Price', 'Total Price']

    response = comp.getVariantsSales()
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

    response = comp.getVariants()
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

def formatPayment(comp):
        response = comp.getPayments()
        pay_df = pd.DataFrame(response.json()['result']['records'])
        to_keep = ('payment_method_id', 'amount', 'pos_order_id')
        pay_df = rm_headers(pay_df, to_keep)
        pay_df['payment_method_id'] = pay_df['payment_method_id'].apply(replace)
        pay_df['pos_order_id'] = pay_df['pos_order_id'].apply(lambda x: replace(x, index=0))
        pay_df = pay_df.loc[pay_df['payment_method_id'].isin(['Cash', 'Transfer', 'POS', 'Customer Account'])]
        t_cash = sum(pay_df.loc[pay_df['payment_method_id'] == 'Cash']['amount'].to_list())
        t_pos = sum(pay_df.loc[pay_df['payment_method_id'] == 'POS']['amount'].to_list())
        t_trfs = sum(pay_df.loc[pay_df['payment_method_id'] == 'Transfer']['amount'].to_list())
        t_cus = sum(pay_df.loc[pay_df['payment_method_id'] == 'Customer Account']['amount'].to_list())
        totals = t_cash, t_trfs, t_pos, t_cus

        response = comp.getOrders()
        order_df = pd.DataFrame(response.json()['result']['records'])
        to_keep = ('partner_id', 'amount_total', 'id')
        order_df = rm_headers(order_df, to_keep)
        order_df = order_df.loc[order_df['partner_id'] != False]
        order_df['partner_id'] = order_df['partner_id'].apply(replace)
        order_df.rename(columns={'id':'pos_order_id', 'partner_id': 'Customer'}, inplace=True)

        df = pd.merge(pay_df, order_df, how='inner', on='pos_order_id')
        df = df.loc[df['payment_method_id'] == 'Customer Account']
        df.drop(columns=['amount_total', 'payment_method_id', 'pos_order_id'], inplace=True)
        df['In'] = df['amount'].apply(lambda x: abs(x) if(x<0) else '')
        df['Out'] = df['amount'].apply(lambda x: x if(x>=0) else '')
        df = df[['Customer', 'In', 'Out']]
        dep_increase = sum([pay if pay != '' else 0 for pay in df['In'].to_list()])
        dep_decrease = sum([pay if pay != '' else 0 for pay in df['Out'].to_list()])
        payment_dict = df.to_dict('records')

        return payment_dict, totals, dep_increase, dep_decrease

def variants_expanded(comp):
        response = comp.getVariants()
        tem_df = pd.DataFrame(response.json()['result']['records'])
        to_keep = ['categ_id', 'name', 'uom_id', 'product_template_variant_value_ids']
        tem_df = rm_headers(df=tem_df, to_keep=to_keep)
        tem_df['categ_id'] = tem_df['categ_id'].apply(replace)
        tem_df['uom_id'] = tem_df['uom_id'].apply(replace)
        tem_df['product_template_variant_value_ids'] = tem_df['product_template_variant_value_ids'].apply((lambda x: replace(cell=x, index=0)))
        template_ids = tem_df['product_template_variant_value_ids'].to_list()
        template_ids = [x for x in template_ids if str(x) != 'nan']
        
        response = comp.attachVariants(template_ids)
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

def getUpdatedQuant(comp, location, categs=None):
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
        var_df = variants_expanded(comp)

        response = comp.getQuants()
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

def get_time():
    today = dt.datetime.today()
    yesterday = dt.datetime.today()-dt.timedelta(1)
    try:
        comp = od.Session(PASSWORD, EMAIL, DOMAIN)
        start_yr, start_month, start_day, start_hour, start_min, start_sec, comp = format_pos(comp)
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
    resp = comp.getCategsList()
    df = pd.DataFrame(resp.json()['result']['records'])
    categs = df['display_name'].to_list()

    return start_yr, start_month, start_day, start_hour, start_min, start_sec, end_yr, end_month, end_day, end_hour, end_min, end_sec, categs, comp

def format_pos(comp):
    response = comp.getPosData()
    if 'error' in response.json():
        if os.path.exists('session_id.txt'):
            os.remove('session_id.txt')
        comp = od.Session(PASSWORD, EMAIL, DOMAIN)
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
    return start_yr, start_month, start_day, hour, min, sec, comp