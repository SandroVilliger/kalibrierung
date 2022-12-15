'''
Programm für die Kalibrierung Kraftsensoren

Programm: Sandro Villiger
'''

# Import der Pakete

import pandas as pd

import numpy as np
import matplotlib.pyplot as plt

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

for x in range(len(Laststufen)):
    Werte_F = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)]
    Werte_F_index = Werte[(Werte['Referenz kN'] >= Laststufen[x] - 0.01) & (Werte['Referenz kN'] <= Laststufen[x] + 0.01)].index.values
    res = np.where(Werte_F_index[:-1] + 1 != Werte_F_index[1:])[0]

    res = res[res > 10]

    X_max_1 = Werte_F[:res[0] + 1]
    X_max_2 = Werte_F[res[0] + 1:res[1] + 1]
    X_max_3 = Werte_F[res[1] + 1:res[2] + 1]
    X_max_4 = Werte_F[res[2] + 1:]
    Mittelwert = X_max_4.mean()
    print(x)
    print(Mittelwert)

    Mittelwerte_Refernz[x].append(X_max_1.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_2.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_3.mean()[1])
    Mittelwerte_Refernz[x].append(X_max_4.mean()[1])

    Mittelwerte_Prüfling[x].append(X_max_1.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_2.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_3.mean()[2])
    Mittelwerte_Prüfling[x].append(X_max_4.mean()[2])

Mittelwerte_Refernz = np.array(Mittelwerte_Refernz)
Mittelwerte_Prüfling = np.array(Mittelwerte_Prüfling)

Refernz_mean = np.mean(Mittelwerte_Refernz, axis=1)
Prüfling_mean = np.mean(Mittelwerte_Prüfling, axis=1)

A = ((Refernz_mean - Prüfling_mean) / Laststufen) * 100

# abweichung Referenznormal
Urn= 0.022


# Abweichung Refernznormal

Umv = 0.05

# Abweichung

U=np.sqrt((1+Urn)*(1+Umv))

print(U)

k=2

text = ['F1', 'F2','F3', 'F4', 'F5', 'F6', 'F7','F8']

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(abs(Laststufen),A,color='green')
plt.plot(abs(Laststufen),A-k*(U-1),color='blue',linestyle='--')
plt.plot(abs(Laststufen),A+k*(U-1),color='blue',linestyle='--')
plt.hlines(0.2,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.hlines(-0.2,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.xticks(abs(Laststufen), text)
plt.grid()
plt.show()

