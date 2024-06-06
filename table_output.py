"""
formats the pandas data frame to be printed in an email
Author: Ryan Morlando
Created: 
Updated: 
V1.0.0
Patch Notes:

To Do:

"""
import pandas as pd
import numpy as np
from os import path
from datetime import date, timedelta
from dotenv import load_dotenv
from os import environ as env
from pathlib import Path
from emailer import send_email
import sys
import re


def zebra_stripe(data):
    s = data.index % 2 != 0
    s = pd.concat([pd.Series(s)] * data.shape[1], axis=1)  # 6 or the n of cols u have
    z = pd.DataFrame(np.where(s, 'background-color:#D1D1D1', ''), index=data.index, columns=data.columns)
    return z


# environment variables
if load_dotenv(r'vars\vars.env') is False:
    with open(r'vars\vars.env', 'w') as file:
        file.write("email=''\nemail_pwd=''\nreceivers=''\nwinner_count='0'")
    exit('Failed to load environment vars. Fill out vars.env')

# for email
receivers = str(env.get('receivers'))
email = str(env.get('email'))
password = str(env.get('email_pwd'))

if receivers == '':
    exit('receivers variable empty. Need at least one receiver email filled out in vars.env')

if email == '':
    exit('email variable empty. Fill out sender email in vars.env')

if password == '':
    exit('email_pwd variable empty. Fill out sender email password in vars.env')

email_credentials = {
    "host": "smtp.gmail.com", "port": "465", "login": email, 'pwd': password, "sender": email
}

receivers = re.split(r',|\s', receivers)

yesterday = str(date.today() - timedelta(days=1))

subject = f'MLB Game Through {yesterday}'


# game files path
file_path = path.dirname(path.abspath(__file__)) + '/game_files/current_game.xlsx'

if not path.exists(file_path):
    exit('No current game to email. Run Updater.py to get started')

# format table
current_game = pd.read_excel(file_path, dtype=str).replace({np.nan: None})

current_game['total'] = current_game['total'].astype('int')

current_game_style = current_game.style \
    .apply(zebra_stripe, axis=None) \
    .set_table_styles([{'selector': 'td',
                        'props': 'border: 2px solid black; border-collapse: collapse'},
                       {'selector': 'thead, th',
                        'props': 'background-color: #D3D3D3; color: black; '
                        'border-right: 2px solid black; '
                        'border-collapse: collapse; '
                        'padding: 20px'}
                       ]) \
    .set_properties(**{'padding': '20px', 'border': '2px solid black', 'border-collapse': 'collapse'})

current_game_style.to_excel(file_path, index=False)

# send email
send_email(credentials=email_credentials, subject=subject, body="See Attached", receivers=receivers,
           attachment_paths=[file_path])
