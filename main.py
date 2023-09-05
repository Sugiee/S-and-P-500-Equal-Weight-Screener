import pandas as pd
import requests                                 # used to make http requests to api to receive data 
import xlsxwriter
import math

from secrets_ import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv('sp_500_stocks.csv')       # current list of stocks on S&P 500 to search     
                                                                                            
# parse the api call
my_columns = ['Ticker', 'Stock Price', 'Market Capitilisation', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)    # pandas dataframe is a 2d structure

# batch API calls to improve performance
def chunks(lst, n):     # used to split list of stocks in to sub-lists
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'], 100))     # 100 elements per sub-list
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))  # element split with ',' in list to separate

for symbol_string in symbol_strings:    # batch request 100 stocks at a time
    batch_api_call_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()      # base url, starts every http request + endpoint + token to let api know im allowed to use
    my_row = []
    for symbol in symbol_string.split(','):
        if symbol == 'DISCA' or symbol == 'HFC' or symbol == 'VIAC' or symbol == 'WLTW': # stocks where took out fund so stocks list outdated, API is unable to get info for these stocks
            continue 
        my_row =\
        [
            symbol,
            data[symbol]['quote']['latestPrice'],
            data[symbol]['quote']['marketCap'],
            'N/A'
        ]
        final_dataframe.loc[len(final_dataframe)] = my_row # add row of current stock data to dataframe

# Calculate the Number of Share to Buy
portfolio_size = input("Enter the value of you portfolio:")

try:
    val = float(portfolio_size)
    print(val)
except ValueError:
    print("Thats not a number! \n Please try again:")
    portfolio_size = input("Enter the value of you portfolio:")
    val = float(portfolio_size)
    
position_size = val/len(final_dataframe.index)
for i in range(0,len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

# write shares to buy to excel sheet
writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter') 
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

#format
background_colour = '#0a0a23'
font_colour = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color' : font_colour,
        'bg_color' : background_colour,
        'border' : 1
    }    
)
pound_format = writer.book.add_format(
    {
        'num_format' : '£0.00',
        'font_color': font_colour,
        'bg_color' : background_colour,
        'border' : 1
    }    
)
integer_format = writer.book.add_format(
    {
        'num_format' : '0',
        'font_color': font_colour,
        'bg_color' : background_colour,
        'border' : 1
    }    
)
column_formats =  {
    'A' :['Ticket', string_format],
    'B' :['Stock Price', pound_format],
    'C' :['Market Capitilisation', pound_format],
    'D' :['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])
    
writer._save()
