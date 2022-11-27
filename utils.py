def convert_to_bundle(qty, uom, categ=None, format='long'):
    def segment(per_bundle, pack_name, unit_name):
        bundle = int(qty//per_bundle)
        rem_yards = qty%per_bundle
        if int(rem_yards) == rem_yards:
            rem_yards = int(rem_yards)
        if rem_yards != 0 and bundle != 0 and per_bundle != 1:
            formatted_qty = f'{bundle} {pack_name} - {rem_yards} {unit_name}'
        elif bundle == 0.0:
            formatted_qty = f'{rem_yards} {unit_name}'
        else:
            formatted_qty = f'{bundle} {pack_name}'
        return formatted_qty

    if format == 'long':
        if categ == 'SEGO':
            qty = segment(1, 'Piece(s)', 'Piece(s)')
        elif categ == 'ASO OKE':
            qty = segment(8, 'Pack(s)', 'Strand(s)')
        elif uom == 'Yards':
            qty = segment(15, 'Bundle(s)', 'Yard(s)')
        else:
            qty = f'{qty} {uom}'
    elif format == 'short':
        if qty>=0:
            if categ == 'SEGO':
                qty = segment(1, 'Pc', 'Pc')
            elif categ == 'ASO OKE':
                qty = segment(8, 'Pck', 'S')
            elif uom == 'Yards':
                qty = segment(15, 'B', 'Y')
            else:
                qty = f'{qty} {uom}'
        else:
            qty = f'{qty} {uom}'
    return qty

def apply_conversion(dict_list, qty_key='Qty', categ_key='Category', uom_key='uom'):
    for dict_ in dict_list:
        dict_[qty_key] = convert_to_bundle(dict_[qty_key], dict_[uom_key], dict_[categ_key])
    return dict_list

def pd_apply_conversion(x):
    return convert_to_bundle(x['product_qty'], x['uom'], x['product_categ_id'], format='short')

def replace(cell, index=1):
    cell = list(cell)
    try:
        cell = cell[index]
    except:
        try:
            cell = cell[0]
        except:
            return None
    return cell

def rm_headers(df, to_keep):
    headers = set(df.columns)
    to_keep = set(to_keep)
    headers.symmetric_difference_update(to_keep)
    df.drop(list(headers), inplace=True, axis=1)
    return df

def rm_other_fields(pro):
    data_list = pro.split(' ')
    for i, data in enumerate(data_list):
        if '(' in data and len(data_list) == 2:
            return pro
        elif len(data_list) == 1:
            return pro
        elif '(' in data and len(data_list) >= 3:
            data = data_list[i-1:i]
            pro = ' '.join(data)
            return pro
        else:
            if i < len(data_list)-1:
                continue
            else:
                pro = data_list[len(data_list)-1]
                return pro


