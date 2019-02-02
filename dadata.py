import openpyxl
import argparse
from dadata import DaDataClient


client = DaDataClient(key='', secret='')


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('file_with_address', help='A file which contains addresses')
    arg = parser.parse_args()
    return arg.file_with_address


def get_fias_code(client_adress):
    client.address = client_adress
    result = client.address.request()
    if result == 200:
        return client.result.fias_code
    else:
        return None


def write_infromation_into_file(my_file):
    work_book = openpyxl.load_workbook(filename=my_file)
    work_sheet = work_book.active
    work_sheet['C1'] = 'Код ФИАС'
    address_column = work_sheet['B']
    for address in range(1, len(address_column)):
        work_sheet.cell(row=address+1, column=3).value = get_fias_code(address_column[address].value)
    work_book.save(my_file)


if __name__ == '__main__':
    file_name = get_args()
    write_infromation_into_file(file_name)
