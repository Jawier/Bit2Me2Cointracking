import os
import sys

import pandas as pd


def extract_dataframe_from_excel_file(excel_file_name):
    return pd.read_excel(excel_file_name, sheet_name='1', index_col=0, skiprows=1)


def transfrom_to_cointracker_format(df):
    df_new = df.rename(columns={'TO AMOUNT': 'Buy Amount', 'TO CURRENCY': 'Buy Currency', 'FROM AMOUNT': 'Sell Amount',
                                'FROM CURRENCY': 'Sell Currency', 'FEE AMOUNT': 'Fee', 'FEE CURRENCY': 'Fee Currency',
                                'EXCHANGE': 'Exchange', 'GROUP': 'Trade-Group', 'DESCRIPTION': 'Comment',
                                'DATE': 'Date', 'ID': 'Tx-ID', 'TO RATE EUR': 'Buy Value in Account Currency',
                                'FROM RATE EUR': 'Sell Value in Account Currency'},

                       index={'Deposit': 'Depósito', 'Trade': 'Operación', 'Reward / Bonus': 'Ingresos por intereses',
                              'Withdrawal': 'Retirada'})

    return df_new[['Buy Amount', 'Buy Currency', 'Sell Amount', 'Sell Currency', 'Fee', 'Fee Currency', 'Exchange',
                   'Trade-Group', 'Comment', 'Date', 'Tx-ID', 'Buy Value in Account Currency',
                   'Sell Value in Account Currency']]


def save_to_csv(dataframe, csv_file_name):
    return dataframe.to_csv(f"{csv_file_name}.csv")


def is_xslx(file_name):
    if os.path.isfile(file_name):
        with open(file_name, 'rb') as f:
            first_four_bytes = f.read()[:4]
            return first_four_bytes == b'PK\x03\x04'
    return False


if __name__ == '__main__':
    excel_file_name = sys.argv[1]
    if not is_xslx(excel_file_name):
        print('The Excel file does not detected')
        exit(1)

    root, ext = os.path.splitext(excel_file_name)
    df_bit2me = extract_dataframe_from_excel_file(excel_file_name)
    df_cointracker = transfrom_to_cointracker_format(df_bit2me)
    save_to_csv(df_cointracker, root)
