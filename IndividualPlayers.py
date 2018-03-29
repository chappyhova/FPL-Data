import pandas as pd
import numpy as np
import re
from xlsxwriter.utility import xl_rowcol_to_cell
from tkinter import *
import json
import time
import aiohttp
import asyncio
import aiofiles

from pip._vendor import requests

if __name__ == '__main__':
    player_link = 'https://fantasy.premierleague.com/drf/element-summary/'
    fplData = requests.get('https://fantasy.premierleague.com/drf/bootstrap-static').json()
    fplData = pd.DataFrame(fplData['elements'])
    print(fplData['id'])
    numberOfPlayers = (len(fplData['id']))
    print(numberOfPlayers)
    URL_LIST = []
    player_dict = {}

for i in range(1, numberOfPlayers+1):
    player_address = player_link+str(i)
    URL_LIST.append(player_address)


async def get_players():
    counter = 1
    async with aiohttp.ClientSession() as session:
        for player in URL_LIST:
            async with session.get(player) as resp:
                data = await resp.json()
                data = pd.DataFrame(data['history'])
                player_id = fplData.loc[fplData['id'] == counter]
                data['first_name'] = player_id['first_name']
                surname = data['second_name'] = player_id['second_name']
                print(surname)
                data['cost'] = player_id['now_cost']
                data['position'] = player_id['element_type']
                data['position'] = data['position'].astype(str)
                data['creativity'] = data['creativity'].astype(float)
                data['threat'] = data['threat'].astype(float)
                player_dict[counter] = data
                print(counter)
                counter += 1
    return player_dict


def sort_players():
    for player, val in player_dict.items():
        val['points'] = val.loc[val['round'] > 10, 'total_points'].sum()
        val['minutes'] = val.loc[val['round'] > 10, 'minutes'].sum()
        val['threat'] = val.loc[val['round'] > 10, 'threat'].sum()
        val['creativity'] = val.loc[val['round'] > 10, 'creativity'].sum()
        player_value = (val['points'] / val['minutes']) / val['cost'] * 10000
        val['value'] = player_value
    return player_dict


loop = asyncio.get_event_loop()
loop.run_until_complete(get_players())
loop.close()

new_player_dict = {}


def not_sure(fpl_dict):
    player = 1
    for item, value in fpl_dict.items():
        value = value.iloc[[0]]
        new_player_dict[player] = value
        player += 1


not_sure(sort_players())
print(new_player_dict)

concatDf = pd.concat(new_player_dict)
concatDf = concatDf[
    ['value', 'first_name', 'second_name', 'minutes', 'cost', 'points', 'bps', 'position', 'threat', 'creativity']]

# Changing element types to positions

concatDf['position'] = concatDf['position'].str.replace('1', 'Goalkeeper').replace('2', 'Defender').replace('3',
                                                                                                            'Midfielder').replace(
    '4', 'Striker')

finalDF = (concatDf.loc[(concatDf['minutes'] > 1) & (concatDf['value'] > 4)])

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
