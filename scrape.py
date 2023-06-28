import pandas as pd
import requests
import os
import matplotlib.pyplot as plt

import matplotlib.font_manager as font_manager

dir = "C:\\Users\\Student\\Desktop"

for font in font_manager.findSystemFonts(dir):
    font_manager.fontManager.addfont(font)

eventnum = ''
eventname = ''

year = int(input('Enter Event Year: '))

if year == 2022:
    eventnum = '140'
    eventname = 'BOF:ET'
elif year == 2021:
    eventnum = '137'
    eventname = 'BOFXVII'
elif year == 2020:
    eventnum = '133'
    eventname = 'BOFXVI'
elif year == 2019:
    eventnum = '127'
    eventname = 'BOFXV'
elif year == 2018:
    eventnum = '123'
    eventname = 'G2R2018'
elif year == 2017:
    eventnum = '116'
    eventname = 'BOFU2017'
elif year == 2016:
    eventnum = '110'
    eventname = 'BOFU2016'
elif year == 2015:
    eventnum = '104'
    eventname = 'BOFU2015'
elif year == 2014:
    eventnum = '96'
    eventname = 'G2R2014'
elif year == 2013:
    eventnum = '88'
    eventname = 'BOF2013'



url = f"https://manbow.nothing.sh/event/event.cgi?action=sp&event={eventnum}"
html = requests.get(url).content.decode('cp932')

tables = pd.read_html(html)
sortedtable = tables[0].sort_values(by=['Total'], ascending=True)

top20 = sortedtable.tail(20)

plt.rcParams["font.family"] = "Noto Sans JP"

total_x = top20['Total'].values.tolist()
total_y = top20['Title'].values.tolist()

plt.barh(total_y, total_x)
plt.title(eventname)
plt.show()

# sortedtable.to_excel("temp.xlsx", index=False)

# filename = f"{eventname}.xlsx"

# os.rename("hello.xlsx", filename)
