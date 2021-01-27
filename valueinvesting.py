from statistics import mean
from scipy.stats import percentileofscore as score
from secrets import IEX_CLOUD_API_TOKEN
import numpy as np  # The Numpy numerical computing library
import pandas as pd  # The Pandas data science library
import requests  # The requests library for HTTP requests in Python
import xlsxwriter  # The XlsxWriter libarary for
import math  # The Python math module
stocks = pd.read_csv('sp_500_stocks.csv')


def portfolio_input():
    global portfolio_size
    portfolio_size = input("Enter the value of your portfolio:")

    try:
        val = float(portfolio_size)
    except ValueError:
        print("That's not a number! \n Try again:")
        portfolio_size = input("Enter the value of your portfolio:")


symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
pe_ratio = data['peRatio']
# Function sourced from
# https://stackoverflow.com/questions/312443/how-do-you-split-a-list-into-evenly-sized-chunks


def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))
#     print(symbol_strings[i])

my_columns = ['Ticker', 'Price',
              'Price-to-Earnings Ratio', 'Number of Shares to Buy']
4  # blank DataFrame as usual
final_dataframe = pd.DataFrame(columns=my_columns)
# for the symbol each append each row with the data from prior
for symbol_string in symbol_strings:
    #     print(symbol_strings)
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series([symbol,
                       data[symbol]['quote']['latestPrice'],
                       data[symbol]['quote']['peRatio'],
                       'N/A'
                       ],
                      index=my_columns),
            ignore_index=True)


# sort the value
final_dataframe.sort_values('Price-to-Earnings Ratio', inplace=True)
# take the value that is more than 0 pe Ratio because negative earning need to be remove
final_dataframe = final_dataframe[final_dataframe['Price-to-Earnings Ratio'] > 0]
# take the top 50
final_dataframe = final_dataframe[:50]
# reset the index
final_dataframe.reset_index(inplace=True)
# drop the index column
final_dataframe.drop('index', axis=1, inplace=True)
portfolio_input()
position_size = float(portfolio_size)/len(final_dataframe.index)
for row in final_dataframe.index:
    final_dataframe.loc[row, 'Number of Shares to Buy'] = math.floor(
        position_size/final_dataframe.loc[row, 'Price'])
symbol = 'AAPL'
batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=advanced-stats,quote&symbols={symbol}&token={IEX_CLOUD_API_TOKEN}'
data = requests.get(batch_api_call_url).json()
# for abbrevation to make it simpler to put value inside the dataframe
# P/E Ratio
pe_ratio = data[symbol]['quote']['peRatio']

# P/B Ratio
pb_ratio = data[symbol]['advanced-stats']['priceToBook']

# P/S Ratio
ps_ratio = data[symbol]['advanced-stats']['priceToSales']

# EV/EBITDA(Enterprise value divided by Earnings Before taxes, depreciation, and amortization)
enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
ebitda = data[symbol]['advanced-stats']['EBITDA']
ev_to_ebitda = enterprise_value/ebitda

# EV/GP enterprise value divided by the gross profit
gross_profit = data[symbol]['advanced-stats']['grossProfit']
ev_to_gross_profit = enterprise_value/gross_profit
rv_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'Price-to-Earnings Ratio',
    'PE Percentile',
    'Price-to-Book Ratio',
    'PB Percentile',
    'Price-to-Sales Ratio',
    'PS Percentile',
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'RV Score'
]

rv_dataframe = pd.DataFrame(columns=rv_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
        ebitda = data[symbol]['advanced-stats']['EBITDA']
        gross_profit = data[symbol]['advanced-stats']['grossProfit']

        try:
            ev_to_ebitda = enterprise_value/ebitda
        except TypeError:
            ev_to_ebitda = np.NaN

        try:
            ev_to_gross_profit = enterprise_value/gross_profit
        except TypeError:
            ev_to_gross_profit = np.NaN

        rv_dataframe = rv_dataframe.append(
            pd.Series([
                symbol,
                data[symbol]['quote']['latestPrice'],
                'N/A',
                data[symbol]['quote']['peRatio'],
                'N/A',
                data[symbol]['advanced-stats']['priceToBook'],
                'N/A',
                data[symbol]['advanced-stats']['priceToSales'],
                'N/A',
                ev_to_ebitda,
                'N/A',
                ev_to_gross_profit,
                'N/A',
                'N/A'
            ],
                index=rv_columns),
            ignore_index=True
        )
rv_dataframe[rv_dataframe.isnull().any(axis=1)]
# this is to create the missing data of the form and fill it with the average value from the data
# use the column inside this only because the other column that contain not an integer will give error because cannot count mean for string value
for column in ['Price-to-Earnings Ratio',
               'Price-to-Book Ratio',
               'Price-to-Sales Ratio', 'EV/EBITDA',
               'EV/GP']:
    # as usual use the inplace= True to make sure the data is saved
    rv_dataframe[column].fillna(rv_dataframe[column].mean(), inplace=True)
rv_dataframe[rv_dataframe.isnull().any(axis=1)]
metrics = {
    'Price-to-Earnings Ratio': 'PE Percentile',
    'Price-to-Book Ratio': 'PB Percentile',
    'Price-to-Sales Ratio': 'PS Percentile',
    'EV/EBITDA': 'EV/EBITDA Percentile',
    'EV/GP': 'EV/GP Percentile'
}

for row in rv_dataframe.index:
    for metric in metrics.keys():
        # same as prior project to get all the percentile value of the column use the percentile of score with the list value count with the value
        rv_dataframe.loc[row, metrics[metric]] = score(
            rv_dataframe[metric], rv_dataframe.loc[row, metric])/100

# Print each percentile score to make sure it was calculated properly
# for metric in metrics.values():
#    print(rv_dataframe[metric])
# loop for each row inside the index
for row in rv_dataframe.index:
    # create an empty list to insert the for lopp with the value from metrics which is the all percentile that has been count
    value_percentiles = []
    # looping for each metric to append the value inside the list
    for metric in metrics.keys():
        # append all the percentile inside the list and then put the mean of the list inside the rv score table and then emptied the list
        value_percentiles.append(rv_dataframe.loc[row, metrics[metric]])
    rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)
# not in descending because wewant to take from the lowest score for the best value
# on this result because we want to take the cheapest stock form the market for value investing
rv_dataframe.sort_values(by='RV Score', inplace=True)
rv_dataframe = rv_dataframe[:50]
rv_dataframe.reset_index(drop=True, inplace=True)
portfolio_input()
position_size = float(portfolio_size) / len(rv_dataframe.index)
for i in range(0, len(rv_dataframe['Ticker'])-1):
    rv_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size / rv_dataframe['Price'][i])
writer = pd.ExcelWriter('value_strategy.xlsx', engine='xlsxwriter')
rv_dataframe.to_excel(writer, sheet_name='Value Strategy', index=False)
background_color = '#0a0a23'
font_color = '#ffffff'

string_template = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_template = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

integer_template = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

float_template = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

percent_template = writer.book.add_format(
    {
        'num_format': '0.0%',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
column_formats = {
    'A': ['Ticker', string_template],
    'B': ['Price', dollar_template],
    'C': ['Number of Shares to Buy', integer_template],
    'D': ['Price-to-Earnings Ratio', float_template],
    'E': ['PE Percentile', percent_template],
    'F': ['Price-to-Book Ratio', float_template],
    'G': ['PB Percentile', percent_template],
    'H': ['Price-to-Sales Ratio', float_template],
    'I': ['PS Percentile', percent_template],
    'J': ['EV/EBITDA', float_template],
    'K': ['EV/EBITDA Percentile', percent_template],
    'L': ['EV/GP', float_template],
    'M': ['EV/GP Percentile', percent_template],
    'N': ['RV Score', percent_template]
}

for column in column_formats.keys():
    writer.sheets['Value Strategy'].set_column(
        f'{column}:{column}', 25, column_formats[column][1])
    writer.sheets['Value Strategy'].write(
        f'{column}1', column_formats[column][0], column_formats[column][1])
writer.save()
