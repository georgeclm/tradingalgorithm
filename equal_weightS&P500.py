import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from secrets import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv('sp_500_stocks.csv')
symbol = "AAPL"
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
print(data['symbol'])
price = data['latestPrice']
market_cap = data['marketCap']
my_columns = ['Ticker', 'Stock Price',
              'Market Capitalization', 'Number of Shares to Buy']
# panda dataframe is a 2d array that list row and column
# it is named series so to create a series to create another column inside the table
final_dataframe = pd.DataFrame(columns=my_columns)
# to append a column
final_dataframe.append(
    pd.Series(
        [
            symbol,
            price,
            market_cap,
            'N/A'
        ],
        # to make sure the index that this symbol,price data input inside the same index as my_columns
        index=my_columns
    ),
    # always put this for no error in pandas
    ignore_index=True
)
final_dataframe = pd.DataFrame(columns=my_columns)
for stock in stocks['Ticker'][:5]:
    api_url = f'https://sandbox.iexapis.com/stable/stock/{stock}/quote/?token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(api_url).json()
    final_dataframe = final_dataframe.append(
        pd.Series([
            stock,
            data['latestPrice'],
            data['marketCap'],
            'N/A'
        ],
            index=my_columns
        ),
        ignore_index=True
    )
print(final_dataframe)
# this function to split a list into a smaller list


def chunks(lst, n):
    # Yield successive n-sized chunked from lst
    for i in range(0, len(lst), n):
        yield lst[i:i+n]


# this is the same as prior to print all the data table but in this one we use batch api so it will run faster for api 1 it iis super slow
# this create a list of panda series for a list contain 100
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    # to join all element in this symbol groups using the "," seperator for a string inside the list
    symbol_strings.append(','.join(symbol_groups[i]))
    # print(symbol_strings[i])
final_dataframe = pd.DataFrame(columns=my_columns)
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series([
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['quote']['marketCap'],
                'N/A'
            ],
                index=my_columns
            ),
            ignore_index=True
        )
print(final_dataframe)
portfolio_size = input('Enter the value of your portofolio: ')
try:
    val = float(portfolio_size)
except ValueError:
    print('Please Input an Integer')
    portfolio_size = input('Enter the value of your portofolio: ')
    val = float(portfolio_size)
position_size = val/len(final_dataframe.index)
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size/final_dataframe.loc[i, 'Stock Price'])
print(final_dataframe)
# to create the table to the xlsx file
writer = pd.ExcelWriter('recommended trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index=False)
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })
dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })
integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })
column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format],
}
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(
        f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(
        f'{column}1', column_formats[column][0], column_formats[column][1])
writer.save()
