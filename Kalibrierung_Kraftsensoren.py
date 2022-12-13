'''
Programm für die Kalibrierung Kraftsensoren

Programm: Sandro Villiger
'''

# Import der Pakete

import pandas as pd
import datetime as dt
import numpy as np
import matplotlib.pyplot as plt





# Pfad zu datei und Import der Datei

file = r'C:\Users\Sandro Villiger\switchdrive\Arbeit IBI\Arbeit Uwe\037 Kallibrirung Killian\Testlauf_2_Kraftgeregelt.XLSX'

# Header Name

name=['Zeit', 'Referenz kN', 'Prüfling kN']

# Import der Daten

data= pd.read_excel(file,header=None,names=name)

# Löschen der unnötigen Daten

delet=np.arange(0,49,1)

Werte=data.drop(delet,axis=0)

# Einlesen der Daten in np.array

Zeit=np.round(np.array(Werte['Zeit'],dtype='float64'),1)
Referenz=np.array(Werte['Referenz kN'],dtype='float64')
Prüfling=np.array(Werte['Prüfling kN'],dtype='float64')

'''
Zugverscuch = 1
Druckversuch = 0
'''

Versuch=0

# maximale Prüfkraft

F_max= 2

if Versuch==0:
    F_max=F_max*-1
else:
    F_max=F_max*1

# Prüfstufen

'''F_12_5=F_max*0.125
F_25=F_max*0.25
F_37_5=F_max*0.375
F_5=F_max*0.5
F_62_5=F_max*0.675   # Noch ändern
F_75=F_max*0.75
F_87_5=F_max*0.875'''

y_Beschriftung=np.linspace(0,F_max,num=5)

Stufen=np.array([0.125,0.25,0.375,0.5,0.675,0.75,0.75,0.875,1])
'''plt.figure(figsize=(8, 5), dpi=300)
plt.plot(Zeit,Referenz,color='black')

plt.hlines(0,min(Zeit),max(Zeit),colors='grey')
plt.vlines(min(Zeit),0,F_max,colors='grey')
plt.title('Versuchsablauf Referenznormal', pad=12,fontsize=12)
plt.xlabel('Zeit [s]' ,fontsize=12,)
plt.ylabel('Kraft [kN]' ,fontsize=12,)
plt.yticks(y_Beschriftung)
plt.show()'''

Laststufen=F_max*Stufen
Mittelwerte_Refernz=[[0],
                    [1],
                    [2],
                   [3],
                   [4],
                   [5],
                   [6],
                   [7],
                   [8] ]

for x in range(len(Laststufen)):
    Werte_F = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)]
    Werte_F_index = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)].index.values
    res = np.where(Werte_F_index[:-1] + 1 != Werte_F_index[1:])[0]

    res=res[res>10]
    print(res)
    X_max_1 = Werte_F[:res[0] + 1]
    X_max_2 = Werte_F[res[0] + 1:res[1] + 1]
    X_max_3 = Werte_F[res[1] + 1:res[2] + 1]
    X_max_4 = Werte_F[res[2] + 1:]
    Mittelwert=X_max_4.mean()
    '''print(Mittelwert)
    print(Mittelwert[1])
    print(Mittelwerte_Refernz[x])'''















