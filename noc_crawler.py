import re
import sys
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import json

import subprocess
from datetime import datetime


def get_info():
    with open("info.json", "r") as file:
        data = json.load(file)
        username = data['username']
        professor_name = data['professor_name']

    return username, professor_name

def get_todays_classes():
    username, professor_name = get_info()

    # login information
    password = sys.argv[1]

    #accessing the site
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(f'https://{quote(username)}:{quote(password)}@noc.insper.edu.br')

    driver.switch_to.frame("menu")

    #find the link and click it
    #link = driver.find_element(By.LINK_TEXT, 'Calendário')
    link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Calendário')))
    link.click()

    # Wait for 'Calendário' to be clickable and click it
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//b[text()='Calendário']"))).click()

    # Wait for 'Calendário Professor' to be clickable and click it
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Calendário Professor'))).click()

    # Switch back to the default content
    driver.switch_to.default_content()
    
    # Switch to the relevant frame ('main' in most cases)
    driver.switch_to.frame("main")

    # Find the text box and input the professor's name
    input_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txNome')))
    input_box.clear()  # clear any existing text
    input_box.send_keys(professor_name)

    button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button2')))
    button.click()

    button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'dgProfessor:_ctl2:_ctl0')))
    button.click()

    table = driver.find_element(By.TAG_NAME, "table") # Could be replaced by a more specific selector
    table_html = table.get_attribute('outerHTML')

    # parse table to pandas dataframe
    dataframes = pd.read_html(table_html)
    df = dataframes[-1]  # take the first table

    new_header = df.iloc[0] #grab the first row for the header
    df = df[1:] #take the data less the header row
    df.columns = new_header #set the header row as the df header

    # switch back to the default content
    driver.switch_to.default_content()

    #close the browser
    driver.quit()

    return df

def parse_datetime_string(s):
    day, month, year = s.split('/')
    day = re.sub('\D', '', day)
    month = re.sub('\D', '', month)
    year = re.sub('\D', '', year)[0:2]
    start_time = re.findall('\d{2}:\d{2}', s)[0]
    end_time = re.findall('\d{2}:\d{2}', s)[1]
    return day, month, year, start_time, end_time

def parse_datetime(d):

    days = []
    months = []
    years = []
    start_times = []
    end_times = []

    for i, r in d.iterrows():
        dt, m, y, st, et = parse_datetime_string(r['Data e Hora'])
        days.append(int(dt))
        months.append(int(m))
        years.append(int(y))
        start_times.append(st)
        end_times.append(et)

    d['day'] = days
    d['month'] = months
    d['year'] = years
    d['start_time'] = start_times
    d['end_time'] = end_times

    #drop the old column
    d.drop(columns=['Data e Hora'], inplace=True)

    return d

def create_calendar_events(df):
    for i, r in df.iterrows():
        day = r['day']
        month = r['month']
        year = r['year']
        start_time = r['start_time']
        end_time = r['end_time']

        description = r['Atividade'] + "-" + r['Titulo'] + " Sala: " + r['Sala']

        # Convert to datetime objects
        start_date = datetime.strptime(f'{day} {month} {year} {start_time}', '%d %m %y %H:%M')
        end_date = datetime.strptime(f'{day} {month} {year} {end_time}', '%d %m %y %H:%M')

        # Format for Applescript
        start_date_str = start_date.strftime('%A, %B %d, %y %I:%M:%M %p')
        end_date_str = end_date.strftime('%A, %B %d, %y %I:%M:%M %p')

        # Applescript command
        cmd = f'''
tell application "Calendar"
    tell calendar "Work"
        make new event at end with properties {{summary:"{description}",start date:date "{start_date_str}",end date:date "{end_date_str}"}}
    end tell
end tell
        '''

        print(cmd)

        # Call the Applescript command
        osa_command = ['osascript', '-e', cmd]
        subprocess.run(osa_command)

def main():
    df = get_todays_classes()
    df = parse_datetime(df)
    create_calendar_events(df)

if __name__ == "__main__":
    main()
