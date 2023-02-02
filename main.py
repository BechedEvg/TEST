import requests
import urllib3
import ssl
from openpyxl import load_workbook
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


def get_legacy_session():
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    ctx.options |= 0x4  # OP_LEGACY_SERVER_CONNECT
    session = requests.session()
    session.mount('https://', CustomHttpAdapter(ctx))
    return session


# Reading a list of elements from a source file.
class Exel_RW:

    def read_exel(file_name, count=-1):
        workbook = pd.read_excel(file_name)
        list_product = []
        for elements in workbook.values:
            list_product.append(list(elements)[:count])
        return list_product

    def write_exel(write_lists, file_name):
        workbook = load_workbook(file_name)
        worksheet = workbook["Sheet"]
        for list_values in write_lists:
            worksheet.append(list_values)
            workbook.save(file_name)
            workbook.close()


def get_html(url):
    html = get_legacy_session().get(url)
    return html


class ProductScraper:

    def __init__(self, input_lists, salary):
        self.input_lists = input_lists
        self.salary = salary


def get_list_analogs(dict_product):
    list_analogs = dict_product["analogs"]



def get_write_lists_product(input_lists):
    for list_product in input_lists:
        list_original_product = list_product[:6]
        vendor_cod = list_product[0]
        dict_product = get_emex_dict_products(vendor_cod)
        list_analogs = get_list_analogs(dict_product)


# Get a list of products from a dictionary.
def get_emex_dict_products(vendor_cod):
    list_dicts_product = []
    url_part1 = "https://emex.ru/api/search/search?detailNum="
    url_part2 = "&locationId=29241&showAll=true"
    url = url_part1 + vendor_cod + url_part2
    dict_product = (get_html(url).json().get('searchResult'))
    return dict_product


# Get a ready-made list of goods for recording.
def get_write_list_products(list_product_elements):
    pass


def main():
    pass
    #input_lists = Exel_RW.read_exel("input.xlsx")


if __name__ == '__main__':
    main()


pd = get_emex_dict_products("13050-0D010")
get_list_analogs(pd)
