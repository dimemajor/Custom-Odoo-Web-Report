
from urllib import response
import requests
from bs4 import BeautifulSoup as bs

class Session():
    def __init__(self, password, email, domain):
        POST_URL = '/web/login'
        VARIANTS_API = '/web/dataset/call_kw/report.pos.order/read_group'
        SEARCH_API = '/web/dataset/search_read'
        REFERER_URL = '/web'
        COLOURS_URL = '/web/dataset/call_kw/product.template.attribute.value/read'
        
        website = domain[7:]
        self.baseUrl = f'{domain}{POST_URL}'
        self.searchUrl = f'{domain}{SEARCH_API}'
        self.variantsUrl = f'{domain}{VARIANTS_API}'
        self.colorUrl = f'{domain}{COLOURS_URL}'
        self.password = password
        self.email = email
        self.headers = {
            'authority': website,
            'accept': '*/*',
            'accept-language': 'en-GB,en;q=0.9,en-US;q=0.8',
            'origin': domain,
            'referer': f'{domain}{REFERER_URL}',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="102", "Microsoft Edge";v="102"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.63 Safari/537.36 Edg/102.0.1245.39',
            }

        try:
            self.session_id = self.loadSessionId()
        except:
            print('new_id')
            self.session_id = self.getSessionId()
        self.cookies = {
            '_ga': 'GA1.2.1463518167.1654244215',
            'tz': 'Africa/Lagos',
            'session_id' : self.session_id,
            'cids': '1',
            'frontend_lang': 'en_US',
            '_gid': 'GA1.2.402550658.1654629813',
        }
    
    def getQuants(self):
        json_data = {
            'id': 12,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'stock.quant',
                'domain': [
                    [
                        'location_id.usage',
                        '=',
                        'internal',
                    ],
                ],
                'fields': [
                    'id',
                    'is_outdated',
                    'sn_duplicated',
                    'tracking',
                    'inventory_quantity_set',
                    'location_id',
                    'storage_category_id',
                    'cyclic_inventory_frequency',
                    'priority',
                    'product_id',
                    'product_categ_id',
                    'lot_id',
                    'package_id',
                    'owner_id',
                    'last_count_date',
                    'available_quantity',
                    'quantity',
                    'product_uom_id',
                    'accounting_date',
                    'inventory_quantity',
                    'inventory_diff_quantity',
                    'inventory_date',
                    'user_id',
                    'company_id',
                ],
                'limit': 10000000,
                'sort': 'location_id ASC, inventory_date ASC, product_id ASC, package_id ASC, lot_id ASC, owner_id ASC',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'params': {
                        'cids': 1,
                        'menu_id': 176,
                        'action': 320,
                        'model': 'stock.quant',
                        'view_type': 'list',
                    },
                    'search_default_internal_loc': 1,
                    'search_default_productgroup': 1,
                    'search_default_locationgroup': 1,
                    'bin_size': True,
                },
            },
        }

        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getVariants(self): #variants list without colours
        json_data = {
            'id': 25,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'product.product',
                'domain': [],
                'fields': [
                    'priority',
                    'default_code',
                    'barcode',
                    'name',
                    'product_template_variant_value_ids',
                    'company_id',
                    'lst_price',
                    'standard_price',
                    'pos_categ_id',
                    'categ_id',
                    'product_tag_ids',
                    'type',
                    'qty_available',
                    'virtual_available',
                    'uom_id',
                    'product_tmpl_id',
                    'active',
                ],
                'limit': 1000000,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'bin_size': True,
                },
            },
        }
        

        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def attachVariants(self, ids): #variant colours
        json_data = {
            'id': 15,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'args': [
                    ids,
                    [
                        'display_name',
                    ],
                ],
                'model': 'product.template.attribute.value',
                'method': 'read',
                'kwargs': {
                    'context': {
                        'lang': 'en_US',
                        'tz': 'Europe/London',
                        'uid': 2,
                        'allowed_company_ids': [
                            1,
                        ],
                    },
                },
            },
        }
        response = requests.post(self.colorUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response
         
    def getSessionId(self):
        with requests.Session() as s:
            r = s.get(self.baseUrl)
            soup = bs(r.content,'html.parser')
            g = soup.head.script.text
            g = g.split('\"')
            data = {
                'csrf_token': g[1],
                'login': self.email,
                'password': self.password,
                'redirect': ''
            }
            r = s.post(self.baseUrl, headers=self.headers, data=data)
            session_id = dict(r.cookies)['session_id']
            try:
                with open('session_id.txt', 'w') as f:
                    f.write(session_id)
            except:
                pass
        return session_id

    def loadSessionId(self):
        #get cookies from session.txt
        with open('session_id.txt','r')as f:
            session_id = f.readline()
        return session_id

    def getPosData(self):
        json_data = {
            'id': 12,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'pos.session',
                'domain': [],
                'fields': [
                    'name',
                    'config_id',
                    'user_id',
                    'start_at',
                    'stop_at',
                    'state',
                ],
                'limit': 80,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'bin_size': True,
                },
            },
        }
        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getCategsList(self):
        json_data = {
            'id': 8,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'product.category',
                'domain': [],
                'fields': [
                    'display_name',
                ],
                'limit': 80,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'params': {
                        'menu_id': 176,
                        'cids': 1,
                        'action': 147,
                        'model': 'product.category',
                        'view_type': 'list',
                    },
                    'bin_size': True,
                },
            },
        }
        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def productMoves(self):
        json_data = {
            'id': 187,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'stock.move.line',
                'domain': [
                    [
                        'state',
                        '=',
                        'done',
                    ],
                ],
                'fields': [
                    'date',
                    'reference',
                    'product_id',
                    'lot_id',
                    'location_id',
                    'location_dest_id',
                    'qty_done',
                    'product_uom_id',
                    'company_id',
                    'state',
                    #'product_category_name', 
                ],
                'limit': 100000,
                'sort': 'date ASC',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'params': {
                        'cids': 1,
                        'menu_id': 176,
                        'action': 331,
                        'model': 'stock.move.line',
                        'view_type': 'list',
                    },
                    'search_default_filter_last_12_months': 0,
                    'search_default_done': 1,
                    'search_default_groupby_product_id': 0,
                    'create': 0,
                    'bin_size': True,
                },
            },
        }
        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getProducts(self):
        json_data = {
            'id': 771,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'product.template',
                'domain': [
                    [
                        'type',
                        'in',
                        [
                            'consu',
                            'product',
                        ],
                    ],
                ],
                'fields': [
                    'id',
                    'product_variant_count',
                    'name',
                    'priority',
                    'default_code',
                    'list_price',
                    'qty_available',
                    'uom_id',
                    'type',
                    'product_variant_ids',
                ],
                'limit': 100000,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'search_default_consumable': 1,
                    'default_detailed_type': 'product',
                    'bin_size': True,
                },
            },
        }
        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

class getData(Session):
    def __init__(self, password, email, domain, start_date, end_date, start_time, end_time):
        self.start_date = start_date
        self.start_time = start_time
        self.end_date = end_date
        self.end_time = end_time
        super().__init__(password, email, domain)

    def getVariantsSales(self):
        json_data = {
            'id': 22,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'args': [],
                'model': 'report.pos.order',
                'method': 'read_group',
                'kwargs': {
                    'context': {
                        'pivot_column_groupby': [
                            'date:month',
                        ],
                        'pivot_row_groupby': [
                            'product_id',
                        ],
                    },
                    'domain': [
                        '&',
                    [
                        'date',
                        '>=',
                        f'{self.start_date} {self.start_time}',
                    ],
                    [
                        'date',
                        '<',
                        f'{self.end_date} {self.end_time}',
                        ],
                    ],
                    'fields': [
                        'product_qty:sum',
                        'average_price:avg',
                        'price_total:sum',
                    ],
                    'groupby': [
                        'product_id',
                        'product_categ_id',
                    ],
                    'lazy': False,
                },
            },
        }
        
        response = requests.post(self.variantsUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getMoves(self):
        json_data = {
            'id': 50,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'stock.move.line',
                'domain': [
                    '&',
                    [
                        'state',
                        '=',
                        'done',
                    ],
                    '&',
                    [
                        'date',
                        '>=',
                        f'{self.start_date} {self.start_time}',
                    ],
                    [
                        'date',
                        '<=',
                        f'{self.end_date} {self.end_time}',
                    ],
                ],
                'fields': [
                    'date',
                    'reference',
                    'product_id',
                    'lot_id',
                    'location_id',
                    'location_dest_id',
                    'qty_done',
                    'product_uom_id',
                    'company_id',
                    'state',
                ],
                'limit': 1000000,
                'sort': 'date ASC',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'params': {
                        'cids': 1,
                        'menu_id': 176,
                        'action': 331,
                        'model': 'stock.move.line',
                        'view_type': 'list',
                    },
                    'search_default_filter_last_12_months': 1,
                    'search_default_done': 1,
                    'search_default_groupby_product_id': 1,
                    'create': 0,
                    'bin_size': True,
                },
            },
        }

        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getOrders(self):
        json_data = {
            'id': 18,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'pos.order',
                'domain': [
                    '&',
                    [
                        'date_order',
                        '>=',
                        f'{self.start_date} {self.start_time}',
                    ],
                    [
                        'date_order',
                        '<=',
                        f'{self.end_date} {self.end_time}',
                    ],
                ],
                'fields': [
                    'currency_id',
                    'name',
                    'session_id',
                    'date_order',
                    'pos_reference',
                    'partner_id',
                    'user_id',
                    'amount_total',
                    'state',
                ],
                'limit': 1000000,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'params': {
                        'menu_id': 219,
                        'cids': 1,
                        'action': 361,
                        'model': 'pos.order',
                        'view_type': 'list',
                    },
                    'bin_size': True,
                },
            },
        }
        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response

    def getPayments(self):
        json_data = {
            'id': 27,
            'jsonrpc': '2.0',
            'method': 'call',
            'params': {
                'model': 'pos.payment',
                'domain': [
                    '&',
                    [
                        'payment_date',
                        '>=',
                        f'{self.start_date} {self.start_time}',
                    ],
                    [
                        'payment_date',
                        '<=',
                        f'{self.end_date} {self.end_time}',
                    ],
                ],
                'fields': [
                    'currency_id',
                    'payment_date',
                    'payment_method_id',
                    'pos_order_id',
                    'amount',
                ],
                'limit': 10000000,
                'sort': '',
                'context': {
                    'lang': 'en_US',
                    'tz': 'Europe/London',
                    'uid': 2,
                    'allowed_company_ids': [
                        1,
                    ],
                    'search_default_group_by_payment_method': 1,
                    'bin_size': True,
                },
            },
        }

        response = requests.post(self.searchUrl, cookies=self.cookies, headers=self.headers, json=json_data)
        return response



            

    










