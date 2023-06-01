"""

Author: Ryan Morlando
Created: 
Updated: 
V1.0.0
Patch Notes:

To Do:

"""
import pandas as pd
import requests
from datetime import date, timedelta
from lxml import html
import numpy as np
from os import path

# how many days back you want to run. Ideally you set it to run everyday pulling yesterdays scores
yesterday = str(date.today() - timedelta(days=1))

# game files path
game_path = path.dirname(path.abspath(__file__)) + '/game_files/'
output_path = (game_path + "current_game.xlsx")

# check whether the game is new or in progress
if path.exists(output_path):
    input_path = output_path
else:
    input_path = (game_path + "template.xlsx")

current_data = pd.read_excel(input_path, dtype=str).replace({np.nan: None})

current_data = current_data.set_index('mapping')

url = f"https://www.espn.com/mlb/scoreboard/_/date/{yesterday.replace('-', '')}"
response = requests.get(url=url)

content = html.fromstring(response.text)

games = content.xpath("//div[@class='Scoreboard__Column flex-auto Scoreboard__Column--1 "
                      "Scoreboard__Column--Score Scoreboard__Column--Score--baseball']")

for i in range(len(games)*2):
    try:
        mapping = content.xpath("(//ul[@class='ScoreboardScoreCell__Competitors']/li/div/div/a/div[@class="
                                "'ScoreCell__TeamName ScoreCell__TeamName--shortDisplayName truncate db'])"
                                f"[{i + 1}]/text()")[0]

        score = content.xpath("(//div[@class='ScoreboardScoreCell_Linescores baseball flex justify-end'])"
                              f"[{i + 1}]/div[1]/text()")[0]

    except IndexError:
        continue

    score = int(score)
    if 0 <= score < 14:
        if current_data.loc[[mapping]][score][0] is None:
            current_data.at[f'{mapping}', score] = f"{yesterday[5:].replace('-', '/')}"

current_data.reset_index(drop=False, inplace=True)
current_data = current_data[['mapping', 'team', 'owner', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]]

writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
current_data.to_excel(writer, index=False)
writer.close()
