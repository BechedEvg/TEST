import os
from operator import itemgetter
from time import sleep
import requests
import urllib3
import ssl
from openpyxl import load_workbook, Workbook
import pandas as pd


class CustomHttpAdapter(requests.adapters.HTTPAdapter):
    # "Transport adapter" that allows us to use custom ssl_context.

    def __init__(self, ssl_context=None, **kwargs):
        self.ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, connections, maxsize, block=False):
        self.poolmanager = urllib3.poolmanager.PoolManager(
            num_pools=connections, maxsize=maxsize,
            block=block, ssl_context=self.ssl_context)


# Reading a list of elements from a source file.
class Exel_RW:

    def read_exel(file_name):
        workbook = pd.read_excel(file_name)
        list_product = []
        for elements in workbook.values:
            list_product.append(list(elements))
        return list_product

    def write_exel(write_lists, file_name, sheet_name=0):
        if file_name not in os.listdir():
            workbook = Workbook()
            workbook.save(file_name)
            workbook.close()
        workbook = load_workbook(file_name)
        if sheet_name == 0:
            worksheet = workbook[workbook.sheetnames[0]]
        elif sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook[sheet_name]
        for list_values in write_lists:
            worksheet.append(list_values)
        workbook.save(file_name)
        workbook.close()


class Analysis:

    def __init__(self, list_data, price_original):
        self.list_data = list_data
        self.price_original = price_original

    def price_check(self):
        pass

    def difference_calculation(self):
        pass


def get_legacy_session():
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    ctx.options |= 0x4  # OP_LEGACY_SERVER_CONNECT
    session = requests.session()
    session.mount('https://', CustomHttpAdapter(ctx))
    return session


def get_html(url):
    html = get_legacy_session().get(url)
    return html


def get_lists_dict_analogs(dict_product):
    lists_dict_analogs_completed = []

    if  "analogs" in dict_product:
        list_dict_analogs = dict_product["analogs"]
        for analog_dict in list_dict_analogs:
            offers = analog_dict['offers']
            for offer in offers:
                lists_dict_analogs_completed.append({
                    "vendor_cod": analog_dict["detailNum"],
                    "make": analog_dict['make'],
                    "name": analog_dict['name'],
                    "price": offer['displayPrice']['value'],
                    "rating": offer['rating2']['rating'],
                    "quantity": offer['quantity'],
                    "delivery": offer['delivery']['value']
                })
    return lists_dict_analogs_completed


def get_lists_product(input_lists):
    write_list = []


    count = len(input_lists)
    for list_product in input_lists[:3]:
        print(list_product)#########################
        sleep(2)
        print(count)################################
        count -= 1##############################################


        list_original_product = list_product
        vendor_cod = str(list_product[-1])

        if vendor_cod not in ["nan", "-"]:

            print(vendor_cod)###################################################

            dict_product = get_emex_dict_products(vendor_cod)
            lists_dict_analogs = get_lists_dict_analogs(dict_product)

            list_analog = []
            for dict_analog in lists_dict_analogs:
                if dict_analog["quantity"] == 1000:
                    dict_analog["quantity"] = "под заказ"
                write_list.append(list_original_product +
                                  [dict_analog['vendor_cod'],
                                   dict_analog['make'],
                                   dict_analog['name'],
                                   dict_analog["price"],
                                   dict_analog["rating"],
                                   dict_analog["quantity"],
                                   dict_analog["delivery"],
                                   f"https://emex.ru/products/{dict_analog['vendor_cod']}/{dict_analog['make']}/29241"])
    return write_list


def get_emex_dict_products(vendor_cod):
    url_part1 = "https://emex.ru/api/search/search?detailNum="
    url_part2 = "&locationId=29241&showAll=true"
    url = url_part1 + vendor_cod + url_part2
    dict_product = (get_html(url).json().get('searchResult'))
    return dict_product


def main():
    input_list = Exel_RW.read_exel("input.xlsx")

    product_lists = get_lists_product(input_list)
    Exel_RW.write_exel(product_lists, "korzina.xlsx")



if __name__ == '__main__':
    main()
    pass
