import pandas as pd
import numpy as np
import re
from xlsxwriter.utility import xl_rowcol_to_cell
from tkinter import *

from pip._vendor import requests

fpl_data = requests.get('https://fantasy.premierleague.com/drf/bootstrap-static').json()
fpl_final = pd.DataFrame(fpl_data['elements'])
fpl_final = fpl_final.rename(columns={'element_type': 'position', 'now_cost': "price"})
fpl_final['position'] = (fpl_final['position']).astype(str)
fpl_final['threat'] = (fpl_final['threat']).astype(float)
fpl_final['creativity'] = (fpl_final['creativity']).astype(float)
fpl_final['points_per_game'] = (fpl_final['points_per_game']).astype(float)
fpl_final['minutes'] = (fpl_final['minutes']).astype(int)
value = (fpl_final['total_points'] / fpl_final['minutes']) / fpl_final['price'] * 10000
bps_pm = (fpl_final['bps'] / fpl_final['minutes']) * 100
fpl_final['bonus_points'] = bps_pm
fpl_final['value'] = value

fpl_final = fpl_final[
    ['value', 'first_name', 'second_name', 'minutes', 'price', 'total_points', 'bonus_points', 'position', 'threat',
     'creativity', 'team' ]]

#top_six = fpl_final['team'].isin([1, 4, 5, 9, 11, 12, 16, 17])

fpl_final = fpl_final.sort_values('value', ascending=False)

# Changing element types to positions

fpl_final['position'] = fpl_final['position'].str.replace('1', 'Goalkeeper').replace('2', 'Defender').replace('3',
                                                                                                              'Midfielder').replace(
    '4', 'Striker')

fpl_final['team'] = (fpl_final['team'].replace([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                                               ['Arsenal', 'Bournemouth', 'Brighton', 'Burnley', 'Chelsea',
                                                'Crystal Palace', 'Everton', 'Huddersfield', 'Leicester', 'Liverpool',
                                                'Man City', 'Man Utd', 'Newcastle', 'Southampton', 'Stoke', 'Swansea',
                                                'Spurs', 'Watford', 'West Brom', 'West Ham']))

results = (fpl_final.loc[(fpl_final['minutes'] > 400) & (fpl_final['value'] > 5)])

writer = pd.ExcelWriter('C:/Users/Chappy/Google Drive/FPLnew.xlsx', engine='xlsxwriter')
results.to_excel(writer, index=False, sheet_name='report')
workbook = writer.book
worksheet = writer.sheets['report']

# Formats for values

price_format = workbook.add_format()
price_format.set_num_format('£##"."#"m"')
value_format = workbook.add_format()
value_format.set_num_format('##.#0')

# Colour formats for cells

green_format = workbook.add_format({'bg_color': '#C6EFCE'})
red_format = workbook.add_format({'bg_color': '#FFC7CE'})
orange_format = workbook.add_format({'bg_color': '#FFEB9C'})

# Getting length of results and storing as a string

length_of_result = (len(results))
length_of_result = (str(length_of_result+1))

# Conditional formatting

worksheet.conditional_format('G2:G' + length_of_result, {'type': 'cell',
                                         'criteria': '>=',
                                         'value': 26,
                                         'format': green_format})

worksheet.conditional_format('G2:G' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 0.1,
                                         'maximum': 19.99,
                                         'format': red_format})

worksheet.conditional_format('G2:G' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 20,
                                         'maximum': 25.99,
                                         'format': orange_format})

worksheet.conditional_format('E2:E' + length_of_result, {'type': 'cell',
                                         'criteria': '>=',
                                         'value': 90,
                                         'format': red_format})

worksheet.conditional_format('E2:E' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 10,
                                         'maximum': 60,
                                         'format': green_format})

worksheet.conditional_format('E2:E' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 61,
                                         'maximum': 89,
                                         'format': orange_format})

worksheet.conditional_format('I2:J' + length_of_result, {'type': 'cell',
                                         'criteria': '>=',
                                         'value': 200,
                                         'format': green_format})

worksheet.conditional_format('I2:J' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 0.0,
                                         'maximum': 99.9,
                                         'format': red_format})

worksheet.conditional_format('I2:J' + length_of_result, {'type': 'cell',
                                         'criteria': 'between',
                                         'minimum': 100,
                                         'maximum': 199.9,
                                         'format': orange_format})

# Setting the column width & formats

worksheet.set_column(0, 0, 14, value_format)
worksheet.set_column(1, 2, 20)
worksheet.set_column(3, 5, 12)
worksheet.set_column(4, 4, 12, price_format)
worksheet.set_column(6, 6, 14, value_format)
worksheet.set_column(7, 7, 14)
worksheet.set_column(8, 9, 8)
worksheet.set_column(10, 10, 8)

workbook.close()
