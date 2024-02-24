import requests
import pandas as pd
from bs4 import BeautifulSoup as BS
import re
import os
import xlsxwriter
import datetime



def extract_data(soup):
    s1 = soup.find('ul', id='top-ratios')
    sh = soup.find('h1', class_='h2 shrink-text')
    s2 = s1.find_all('li')

    dic = {'Company Name': sh.text}
    num = [line.text.replace('\n', '') for line in s2]
    cleaned_data = [' '.join(line.split()) for line in num]

    for line in cleaned_data:
        match = re.match(r"([A-Za-z/ ]+)([₹0-9.,\-%/Cr ]+)", line)
        if match:
            dic[match.group(1).strip()] = match.group(2).strip()

    df = pd.DataFrame(dic, index=[0])
    current_date = datetime.datetime.now()
    df['current_date'] = str(current_date)
    return df


def save_to_excel(df, file_name, sheet_name):

    if os.path.isfile(file_name):
        try:
            existing_df = pd.read_excel(
                file_name, sheet_name=sheet_name, engine='openpyxl')
            updated_df = pd.concat([existing_df, df], ignore_index=True)

            with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                updated_df.to_excel(
                    writer, sheet_name=sheet_name, index=False, header=True)
                

        except:
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name,
                            index=False, header=True)
                
    else:
        with pd.ExcelWriter(file_name, engine='xlsxwriter', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet_name,
                        index=False, header=True)
            

url = input('PLEAE PROVIDE ME THE LINK WITH COMPANY NAME FROM SCREENER.IN SO THAT YOU CAN GET SUMMARY:  ')
r = requests.get(url)
soup = BS(r.content, 'html.parser')
df = extract_data(soup)
file_name = 'screnner_summary.xlsx'
sheet_name = str(df['Company Name'].iloc[0].upper())
save_to_excel(df, file_name, sheet_name)
