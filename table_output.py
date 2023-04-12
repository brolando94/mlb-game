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
from datetime import date
from dod_emailer import emailer as email
from dotenv import load_dotenv
from os import environ as env
from pathlib import Path


def zebra_stripe(data):
    s = data.index % 2 != 0
    s = pd.concat([pd.Series(s)] * data.shape[1], axis=1)  # 6 or the n of cols u have
    z = pd.DataFrame(np.where(s, 'background-color:#D1D1D1', ''), index=data.index, columns=data.columns)
    return z


# environment variables
load_dotenv(fr'{Path.home()}\vars.env')

email_credentials = {
    "host": f"{env.get('email_host')}", "port": f"{env.get('email_port')}", "login": f"{env.get('email')}",
    'pwd': f"{env.get('email_pwd')}", "sender": f"{env.get('email_sender')}"
}

# email information
subject = f'MLB Game {date.today()}'
receivers = ['caps1394@gmail.com']

# game files path
file_path = path.dirname(path.abspath(__file__)) + '/game_files/current_game.xlsx'

if not path.exists(file_path):
    exit('No current game to format')

current_game = pd.read_excel(file_path, dtype=str).replace({np.nan: None})

current_game_style = current_game.style \
    .set_table_styles([{'selector': ' , td',
                        'props': 'border: 2px solid black; border-collapse: collapse'},
                       {'selector': 'thead, th',
                        'props': 'background-color: black; color: white; '
                        'border-right: 1px solid white; '
                        'border-collapse: collapse; '
                        'padding: 5px'}
                       ]) \
    .apply(zebra_stripe, axis=None)\
    .set_properties(padding='5px')

current_game_html = current_game_style.hide(axis='index').to_html().replace('None', '')

# body for email
html = f"""\
            <html>
                <head>Hello,<br></br>\nSee Below or Attached Excel File\n </head>
                    <body>
                        <div>
                            <br></br>
                            <b>\nCurrent Game:</b>
                            {current_game_html}
                            <br></br>
                        </div>
                    </body>
                </html>
        """

# send email
email.send_email(subject=subject, body=html, body_type='html', credentials=email_credentials,
                 receivers=receivers, attachment_paths=[file_path]
                 )
