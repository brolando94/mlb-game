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
from dotenv import load_dotenv
from os import environ as env
from pathlib import Path
from emailer import send_email

# environment variables
if load_dotenv(r'vars\vars.env') is False:
    with open(r'vars\vars.env', 'w') as file:
        file.write("email=''\nemail_pwd=''\nreceivers=''")
    exit('Failed to load environment vars. Fill out vars.env')
email = str(env.get('email'))
email_credentials = {
    "host": "smtp.gmail.com", "port": "465", "login": email,
    'pwd': f"{env.get('email_pwd')}", "sender": email
}

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

user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " \
             "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

base_headers = {
    'user-agent': f'{user_agent}',
    'accept': 'text/html,*/*',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'en-US,en;q=0.9',
}

url = f"https://www.espn.com/mlb/scoreboard/_/date/{yesterday.replace('-', '')}"
response = requests.get(url=url, headers=base_headers)
if response.status_code != 200:
    print(f"Error on {yesterday}")
    subject = f"MLB Update Error"
    body = f"Failed to update {yesterday}\n\t{response.status_code}"
    send_email(credentials=email_credentials, body=body, receivers=[email], subject=subject)
    exit()
content = html.fromstring(response.text)

games = content.xpath("//div[@class='Scoreboard__Column flex-auto Scoreboard__Column--1 "
                      "Scoreboard__Column--Score Scoreboard__Column--Score--baseball']")

for i in range(len(games)*2):
    # need to check to make sure that the game actually finished.
    # if you pull to early postponed games that have score will enter the game.
    # on espn the game should go away from the previous day when it restarts the next day
    # but this is a way to avoid that issue. Not clean but works
    outcome = str(content.xpath(f"(//ul[@class='ScoreboardScoreCell__Competitors']/li)[{i + 1}]")[0].attrib)
    if 'winner' not in outcome and 'loser' not in outcome:
        continue

    mapping = content.xpath("(//ul[@class='ScoreboardScoreCell__Competitors']/li/div/div/a/div[@class="
                            "'ScoreCell__TeamName ScoreCell__TeamName--shortDisplayName truncate db'])"
                            f"[{i + 1}]/text()")[0]

    score = content.xpath("(//div[@class='ScoreboardScoreCell_Linescores baseball flex justify-end'])"
                          f"[{i + 1}]/div[1]/text()")[0]

    score = int(score)
    if 0 <= score < 14:
        if current_data.loc[mapping][score] is None:
            current_data.at[f'{mapping}', score] = f"{yesterday[5:].replace('-', '/')}"

current_data.reset_index(drop=False, inplace=True)
current_data = current_data[['mapping', 'team', 'owner', 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]]

# populate totals
total_df = pd.DataFrame()
winner = False
for index, row in current_data.iterrows():
    count = 0
    for i in range(14):
        if row[i] is not None:
            count += 1

    # winner detection
    if count == 14:
        winner = True

    total_df = pd.concat([total_df, pd.DataFrame([[
        row['mapping'],
        count
    ]], columns=['mapping', 'total'])])
current_data = current_data.merge(total_df, how='left', on='mapping')

# write to excel
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
current_data.to_excel(writer, index=False)
writer.close()
# send winner email
if winner:
    send_email(credentials=email_credentials, receivers=[email], subject='Baseball Winner detected', body='')
