import pandas as pd
import re


df = pd.read_excel(
                    '/home/hymar/Pulpit/projekty/docu_downloader/odczynniki.xls',
                    engine='xlrd'
                    )

dostawcy = {'thermo':'https://www.thermofisher.com/pl/en/home.html',
            'sigma':'https://www.sigmaaldrich.com/PL/pl',
            'vwr':'https://pl.vwr.com/store/'}

#funkcja 'refex' wyszukuje klucze 'keys' ze słownika 'dostawcy' iterując po data frame 

def refex(key, link):

    for index, row in df.iterrows():
        if len(re.findall(key,row['web']))==1:
            return [link, row['cat nr'], row['lot nr']]
        
#pętla iterując, wsadza elementy listy 'dostawca' do funkcji 'refex'
for item in dostawcy.items():
    print(refex(item[0], item[1]))