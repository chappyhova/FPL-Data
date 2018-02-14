import csv

import pandas as pd
from pip._vendor import requests

id_list = []
results_dict = {}
gameweeks_required = [18, 19, 20, 21, 22]

player_ids = requests.get("https://fantasy.premierleague.com/drf/leagues-classic-standings/57289?phase=1&le-page=1&ls-page=1").json()
player_ids_2 = requests.get("https://fantasy.premierleague.com/drf/leagues-classic-standings/57289?phase=1&le-page=1&ls-page=2").json()
player_ids = player_ids['standings']['results']
player_ids = player_ids + player_ids_2['standings']['results']

for player in player_ids:
    id_list.append(player['entry'])

print(len(id_list))

for player_id in id_list:
    total_points = 0
    for week in gameweeks_required:
        week_data = requests.get('https://fantasy.premierleague.com/drf/entry/'+str(player_id)+'/event/'+str(week)+'/picks').json()
        new_points = week_data['entry_history']['points']
        transfer_cost = week_data['entry_history']['event_transfers_cost']
        total_points = (new_points + total_points) - transfer_cost
    team_name = requests.get('https://fantasy.premierleague.com/drf/entry/' + str(player_id)).json()
    results_dict[team_name['entry']['name']] = total_points

with open('fplMonthly.csv', 'w', newline="") as f:
    writer = csv.writer(f)
    for row in results_dict.items():
        writer.writerow(row)

