"""
Programm für die Kalibrierung Kraftsensoren

Programm: Sandro Villiger
"""

#%% Pakete


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import PIL as p


#%%



# Pfad zu datei und Import der Datei

file = 'Testlauf_2_Kraftgeregelt.XLSX'

# Header Name

name = ['Zeit', 'Referenz kN', 'Prüfling kN']

# Import der Daten

data = pd.read_excel(file, header=None, names=name)

# Löschen der unnötigen Daten

delet = np.arange(0, 49, 1)


Werte = data.drop(delet, axis=0)

# Einlesen der Daten in np.array

Zeit = np.round(np.array(Werte['Zeit'], dtype='float64'), 1)
Referenz = np.array(Werte['Referenz kN'], dtype='float64')
Prüfling = np.array(Werte['Prüfling kN'], dtype='float64')

'''
Zugverscuch = 1
Druckversuch = 0
'''

Versuch = 0

# maximale Prüfkraft

F_max = 2

if Versuch == 0:
    F_max = F_max * -1
else:
    F_max = F_max * 1

# Messstrom

S=2

# Prüfstufen


y_Beschriftung = np.linspace(0, F_max, num=5)
Stufen = np.array([0.125, 0.25, 0.375, 0.5, 0.675, 0.75, 0.875, 1])

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(Zeit,Referenz,color='black')

plt.hlines(0,min(Zeit),max(Zeit),colors='grey')
plt.vlines(min(Zeit),0,F_max,colors='grey')
plt.title('Versuchsablauf Referenznormal', pad=12,fontsize=12)
plt.xlabel('Zeit [s]' ,fontsize=12,)
plt.ylabel('Kraft [kN]' ,fontsize=12,)
plt.yticks(y_Beschriftung)
plt.show()

Laststufen = F_max * Stufen
Mittelwerte_Refernz = [[], [], [], [], [], [], [],  []]

Mittelwerte_Prüfling = [[], [], [], [], [], [], [], []]

Null_Last=Werte_F = Werte[(Werte['Referenz kN'] >= 0- 0.01) & (Werte['Referenz kN'] <=0 + 0.01)]['Referenz kN'].mean()

for x in range(len(Laststufen)):
    Werte_F = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)]
    Werte_F_index = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)].index.values
    res = np.where(Werte_F_index[:-1] + 1 != Werte_F_index[1:])[0]
    #print(res)
    res = res[res > 10]

    X_max_1 = Werte_F[10:res[0] -9]
    X_max_2 = Werte_F[res[0] + 10:res[1] -9]
    X_max_3 = Werte_F[res[1] + 10:res[2] -9]
    X_max_4 = Werte_F[res[2] + 10:-10]
    Mittelwert = X_max_4.mean()
    #print(x)
    #print(res)

    Mittelwerte_Refernz[x].append(X_max_1.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_2.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_3.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_4.mean()[1])

    Mittelwerte_Prüfling[x].append(X_max_1.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_2.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_3.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_4.mean()[2])
    '''plt.figure(figsize=(8, 5), dpi=300)
    plt.plot(Zeit, Referenz, color='black')
    plt.plot(X_max_1['Zeit'], X_max_1['Referenz kN'], color='blue')
    plt.plot(X_max_2['Zeit'], X_max_2['Referenz kN'], color='red')
    plt.plot(X_max_3['Zeit'], X_max_3['Referenz kN'], color='blue')
    plt.plot(X_max_4['Zeit'], X_max_4['Referenz kN'], color='red')
    plt.vlines(min(Zeit), 0, F_max, colors='grey')
    plt.title('Versuchsablauf Referenznormal', pad=12, fontsize=12)
    plt.xlabel('Zeit [s]', fontsize=12, )
    plt.ylabel('Kraft [kN]', fontsize=12, )
    plt.yticks(y_Beschriftung)
    plt.show()'''
Mittelwerte_Refernz = np.array(Mittelwerte_Refernz)
Mittelwerte_Prüfling = np.array(Mittelwerte_Prüfling)


Refernz_mean = np.mean(Mittelwerte_Refernz, axis=1)
Prüfling_mean = np.mean(Mittelwerte_Prüfling, axis=1)

A = ((Prüfling_mean-Refernz_mean ) / Laststufen) * 100

# abweichung Referenznormal
Urn= 0.022


# Abweichung Messverstärker

Umv = 0.05

# Abweichung

U=np.sqrt((Urn)**2+(Umv)**2)

#print(U)

k=2

text = ['F1', 'F2','F3', 'F4', 'F5', 'F6', 'F7','F8']

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(abs(Laststufen),A,color='green')
plt.plot(abs(Laststufen),A-k*(U),color='blue',linestyle='--')
plt.plot(abs(Laststufen),A+k*(U),color='blue',linestyle='--')
plt.hlines(0.2,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.hlines(-0.2,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.xticks(abs(Laststufen), text)
plt.grid()
plt.show()


'''
Ausgabe für Fehler korrekter 
'''

Korrektur_max= (S/ ((F_max*1000)-(A[-1]/100)*F_max*1000))*F_max*1000

Korrektur_unten=((A[0]/100)*F_max*1000)+(F_max*1000*Stufen[0])

Korrektur_min=(Stufen[0]*S*F_max*1000/Korrektur_unten)-2

print(0,'N','=', Korrektur_min,'mV/V')
print(-F_max*1000,'N','=', Korrektur_max,'mV/V')


