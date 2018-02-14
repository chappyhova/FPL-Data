import pandas as pd
import numpy as np
import re
from xlsxwriter.utility import xl_rowcol_to_cell
from tkinter import *
import json

from pip._vendor import requests

fplData = requests.get('https://fantasy.premierleague.com/drf/bootstrap-static').json()
fplData = pd.DataFrame(fplData['elements'])
numberOfPlayers = (len(fplData['id']))

playerDict = {}
player = 1
while player <= numberOfPlayers:
    playerJson = requests.get('https://fantasy.premierleague.com/drf/element-summary/' + str(player)).json()
    print(player)
    playerJson = pd.DataFrame(playerJson['history'])
    playerJson['first_name'] = (fplData.loc[player - 1]['first_name'])
    playerJson['second_name'] = (fplData.loc[player - 1]['second_name'])
    playerJson['cost'] = (fplData.loc[player - 1]['now_cost'])
    playerJson['position'] = (fplData.loc[player - 1]['element_type'])
    playerJson['position'] = playerJson['position'].astype(str)
    playerJson['creativity'] = playerJson['creativity'].astype(float)
    playerJson['threat'] = playerJson['threat'].astype(float)
    playerDict[player] = playerJson
    player += 1

for item, value in playerDict.items():
    value['points'] = value.loc[value['round'] > 20, 'total_points'].sum()
    value['minutes'] = value.loc[value['round'] > 20, 'minutes'].sum()
    value['threat'] = value.loc[value['round'] > 20, 'threat'].sum()
    value['creativity'] = value.loc[value['round'] > 20, 'creativity'].sum()
    playerValue = (value['points'] / value['minutes']) / value['cost'] * 10000
    value['value'] = playerValue

newPlayerDict = {}
player2 = 1
for item, value in playerDict.items():
    value = value.iloc[[0]]
    newPlayerDict[player2] = value
    player2 += 1

concatDf = pd.concat(newPlayerDict)
concatDf = concatDf[
    ['value', 'first_name', 'second_name', 'minutes', 'cost', 'points', 'bps', 'position', 'threat', 'creativity']]

# Changing element types to positions

concatDf['position'] = concatDf['position'].str.replace('1', 'Goalkeeper').replace('2', 'Defender').replace('3',
                                                                                                              'Midfielder').replace(
        '4', 'Striker')

finalDF = (concatDf.loc[(concatDf['minutes'] > 300) & (concatDf['value'] > 7)])

writer = pd.ExcelWriter('C:/Users/chapp/Google Drive/FPLform.xlsx', engine='xlsxwriter')
finalDF.to_excel(writer, index=False, sheet_name='report')
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

length_of_result = numberOfPlayers
length_of_result = (str(length_of_result + 1))

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

workbook.close()
