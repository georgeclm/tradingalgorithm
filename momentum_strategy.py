from statistics import mean
from secrets import IEX_CLOUD_API_TOKEN
import numpy as np
import pandas as pd
import requests
import math
from scipy.stats import percentileofscore as score
import xlsxwriter
stocks = pd.read_csv('sp_500_stocks.csv')
symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/stats?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
print(data['year1ChangePercent'])
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
              'One-Year Price Return', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)
for symbol_string in symbol_strings:
    # so to get the inside value is in types of data that you want to return in here the stats to take the 1year price and the price to take the current price
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=price,stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['price'],
                    data[symbol]['stats']['year1ChangePercent'],
                    'N/A'

                ],
                index=my_columns),
            ignore_index=True
        )
# use this inplace value set to true to make the final_dataframe get updated not just in this table
final_dataframe.sort_values('One-Year Price Return',
                            ascending=False, inplace=True)
# to modify the data that has been sorted by the one year price return to take only 50 best
final_dataframe = final_dataframe[:50]
# one time action to reset the index
final_dataframe.reset_index(inplace=True)
final_dataframe


def portofolio_input():
    global portofolio_size
    portofolio_size = input('Enter the size of your portofolio: ')
    try:
        float(portofolio_size)
    except ValueError:
        print('That is not a number')
        portofolio_size = input('Enter the size of your portofolio: ')


portofolio_input()
position_size = float(portofolio_size)/len(final_dataframe.index)
for i in range(0, len(final_dataframe)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size/final_dataframe.loc[i, 'Price'])
hqm_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'One-Year Price Return',
    'One-Year Return Percentile',
    'Six-Month Price Return',
    'Six-Month Return Percentile',
    'Three-Month Price Return',
    'Three-Month Return Percentile',
    'One-Month Price Return',
    'One-Month Return Percentile',
    'HQM Score'
]
hqm_dataframe = pd.DataFrame(columns=hqm_columns)
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=price,stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = hqm_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['price'],
                    'N/A',
                    data[symbol]['stats']['year1ChangePercent'],
                    'N/A',
                    data[symbol]['stats']['month6ChangePercent'],
                    'N/A',
                    data[symbol]['stats']['month3ChangePercent'],
                    'N/A',
                    data[symbol]['stats']['month1ChangePercent'],
                    'N/A',
                    'N/A',
                ],
                index=hqm_columns),
            ignore_index=True
        )
# add time periods so each loop from the dataframe will loop again inside the time period so it will count on 1 index all the return percentile
time_periods = [
    'One-Year',
    'Six-Month',
    'Three-Month',
    'One-Month'
]
# create the loops here for each row inside the hqm and take the return percentile in each time period so 1 row will count all the percentile
for row in hqm_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        if hqm_dataframe.loc[row, change_col] == None:
            hqm_dataframe.loc[row, change_col] = 0.0
for row in hqm_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_dataframe.loc[row, percentile_col] = score(
            hqm_dataframe[change_col], hqm_dataframe.loc[row, change_col])/100
# to loop all the rown inside dataframe
for row in hqm_dataframe.index:
    # create the momentum list and it is going to be emptied for each row so after a row got put inside the hqm score then emptied again and loop until the last row
    momentum_percentiles = []
    for time_period in time_periods:
        # take the dataframe return percentile for each line and put it inside the list to calculate the mean
        momentum_percentiles.append(
            hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    # after have all the value inside the each loc put it inside that row the hqm score which is the mean from 1 year to 1 month retun percentile
    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
# same as prior sort the data based on the hqm score
hqm_dataframe.sort_values('HQM Score', ascending=False, inplace=True)
# take the top 50 value and store to the dataframe
hqm_dataframe = hqm_dataframe[:50]
# one time run to reset the index
hqm_dataframe.reset_index(inplace=True, drop=True)
portofolio_input()
position_size = float(portofolio_size) / len(hqm_dataframe.index)
for i in hqm_dataframe.index:
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size/hqm_dataframe.loc[i, 'Price'])
writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, sheet_name="Momentum Strategy", index=False)
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
    'D': ['One-Year Price Return', percent_template],
    'E': ['One-Year Return Percentile', percent_template],
    'F': ['Six-Month Price Return', percent_template],
    'G': ['Six-Month Return Percentile', percent_template],
    'H': ['Three-Month Price Return', percent_template],
    'I': ['Three-Month Return Percentile', percent_template],
    'J': ['One-Month Price Return', percent_template],
    'K': ['One-Month Return Percentile', percent_template],
    'L': ['HQM Score', percent_template]
}
for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(
        f'{column}:{column}', 25, column_formats[column][1])
    writer.sheets['Momentum Strategy'].write(
        f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()
