import numpy as np
import pandas as pd
import requests
import math
from scipy.stats import percentileofscore as score
import xlsxwriter
from statistics import mean

stocks = pd.read_csv('trylq45.csv')
symbol_strings = list(stocks['Ticker'])
oneyearPrices = list(stocks['One-Year Price Return'])
sixmonthPrices = list(stocks['Six-Month Price Return'])
threemonthPrices = list(stocks['Three-Month Price Return'])
onemonthPrices = list(stocks['One-Month Price Return'])

prices = list(stocks['Price'])


def portofolio_input():
    global portofolio_size
    portofolio_size = input('Enter the size of your portofolio: ')
    try:
        float(portofolio_size)
    except ValueError:
        print('That is not a number')
        portofolio_size = input('Enter the size of your portofolio: ')

# my_columns = ['Ticker', 'Price',
#               'One-Year Price Return', 'Number of Shares to Buy']
# final_dataframe = pd.DataFrame(columns=my_columns)
# for i in range(0, len(stocks)):
#     final_dataframe = final_dataframe.append(
#         pd.Series(
#             [
#                 symbol_strings[i],
#                 prices[i]*100,
#                 oneyearPrices[i],
#                 'N/A'
#             ],
#             index=my_columns), ignore_index=True
#     )
# final_dataframe.sort_values('One-Year Price Return',
#                             ascending=False, inplace=True)
# final_dataframe = final_dataframe[:10]
# final_dataframe.reset_index(inplace=True, drop=True)


# portofolio_input()
# position_size = float(portofolio_size)/len(final_dataframe.index)
# for i in range(0, len(final_dataframe)):
#     final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
#         position_size/final_dataframe.loc[i, 'Price'])
# print(final_dataframe)
# now for the big partt using hqm by taking the 1 month, 3 month, 6 month, 1 year price return for more realistic result
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
for i in range(0, len(stocks)):
    hqm_dataframe = hqm_dataframe.append(
        pd.Series(
            [
                symbol_strings[i],
                prices[i]*100,
                'N/A',
                oneyearPrices[i]/100,
                'N/A',
                sixmonthPrices[i]/100,
                'N/A',
                threemonthPrices[i]/100,
                'N/A',
                onemonthPrices[i]/100,
                'N/A',
                'N/A'
            ],
            index=hqm_columns), ignore_index=True
    )
time_periods = [
    'One-Year',
    'Six-Month',
    'Three-Month',
    'One-Month'
]
for row in hqm_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_dataframe.loc[row, percentile_col] = score(
            hqm_dataframe[change_col], hqm_dataframe.loc[row, change_col])/100
for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(
            hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
hqm_dataframe.sort_values('HQM Score', ascending=False, inplace=True)
hqm_dataframe = hqm_dataframe
hqm_dataframe.reset_index(inplace=True, drop=True)
portofolio_input()
position_size = float(portofolio_size) / len(hqm_dataframe.index)
for i in hqm_dataframe.index:
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(
        position_size/hqm_dataframe.loc[i, 'Price'])
writer = pd.ExcelWriter(
    f'momenntum_26-Jan-2021_LQ45_AllRank{portofolio_size}.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, sheet_name="Momentum LQ45", index=False)
background_color = '#000000'
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
        'num_format': 'Rp .00',
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
    writer.sheets['Momentum LQ45'].set_column(
        f'{column}:{column}', 25, column_formats[column][1])
    writer.sheets['Momentum LQ45'].write(
        f'{column}1', column_formats[column][0], column_formats[column][1])
writer.save()
