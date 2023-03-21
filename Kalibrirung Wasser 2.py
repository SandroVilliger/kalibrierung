"""
Programm für die Kalibrierung Wasser

Programm: Sandro Villiger
"""

#%% Pakete


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


#%%

# Pfad zu datei und Import der Datei

file = r'C:\Users\Sandro Villiger\switchdrive\Arbeit IBI\Arbeit Uwe\037 Kallibrirung Killian\Kallibrirung der Wassermenge\2022-12-22-13_45_53.485_Kalibration_Wasser_2.csv'

#%%
# Import der Daten

data = pd.read_csv(file,sep=';')

Referenz=np.array(data['MW'])

Messung=np.array(data['Qw 20'])

Punkte= np.arange(0,len(Referenz),1)

#%%

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(Punkte,Referenz,color='black',label='Referenz')
plt.plot(Punkte,Referenz,color='blue',label='Messung')
plt.grid()
plt.xticks(np.arange(0,10001,1000))
plt.legend()
plt.show()

#%%




Werte_1_r=Referenz[1750:1950]
Werte_2_r=Referenz[2200:2400]
Werte_3_r=Referenz[2600:2900]
Werte_4_r=Referenz[3150:3420]
Werte_5_r=Referenz[3680:4000]
Werte_6_r=Referenz[4220:4500]
Werte_7_r=Referenz[4800:5050]
Werte_8_r=Referenz[5350:5600]
Werte_9_r=Referenz[5940:6300]
Werte_10_r=Referenz[6580:6720]
Werte_11_r=Referenz[6950:7250]
Werte_12_r=Referenz[7500:7800]
Werte_13_r=Referenz[8000:8320]
Werte_14_r=Referenz[8550:8900]
Werte_15_r=Referenz[9100:9400]
Werte_16_r=Referenz[9650:10000]
Werte_17_r=Referenz[10220:10520]

Werte_1_m=Messung[1750:1950]
Werte_2_m=Messung[2200:2400]
Werte_3_m=Messung[2600:2900]
Werte_4_m=Messung[3150:3420]
Werte_5_m=Messung[3680:4000]
Werte_6_m=Messung[4220:4500]
Werte_7_m=Messung[4800:5050]
Werte_8_m=Messung[5350:5600]
Werte_9_m=Messung[5940:6300]
Werte_10_m=Messung[6580:6720]
Werte_11_m=Messung[6950:7250]
Werte_12_m=Messung[7500:7800]
Werte_13_m=Messung[8000:8320]
Werte_14_m=Messung[8550:8900]
Werte_15_m=Messung[9100:9400]
Werte_16_m=Messung[9650:10000]
Werte_17_m=Messung[10220:10520]





#%%


plt.figure(figsize=(8, 5), dpi=300)
plt.plot(Punkte,Referenz,color='black',label='Referenz')
plt.plot(Punkte,Referenz,color='blue',label='Messung')


plt.plot(np.arange(1750,1950),Werte_1_r,color='red',label='Auswertung')
plt.plot(np.arange(2200,2400),Werte_2_r,color='red')
plt.plot(np.arange(2600,2900),Werte_3_r,color='red')
plt.plot(np.arange(3150,3420),Werte_4_r,color='red')
plt.plot(np.arange(3680,4000),Werte_5_r,color='red')
plt.plot(np.arange(4220,4500),Werte_6_r,color='red')
plt.plot(np.arange(4800,5050),Werte_7_r,color='red')
plt.plot(np.arange(5350,5600),Werte_8_r,color='red')
plt.plot(np.arange(5940,6300),Werte_9_r,color='red')
plt.plot(np.arange(6580,6720),Werte_10_r,color='red')
plt.plot(np.arange(6950,7250),Werte_11_r,color='red')
plt.plot(np.arange(7500,7800),Werte_12_r,color='red')
plt.plot(np.arange(8000,8320),Werte_13_r,color='red')
plt.plot(np.arange(8550,8900),Werte_14_r,color='red')
plt.plot(np.arange(9100,9400),Werte_15_r,color='red')
plt.plot(np.arange(9650,10000),Werte_16_r,color='red')
plt.plot(np.arange(10220,10520),Werte_17_r,color='red')

plt.grid()
plt.xticks(np.arange(0,10001,1000))
plt.legend()
plt.show()


#%% Berechnung der Mittelwerte


x1_R=Werte_1_r.mean()
x2_R=Werte_2_r.mean()
x3_R=Werte_3_r.mean()
x4_R=Werte_4_r.mean()
x5_R=Werte_5_r.mean()
x6_R=Werte_6_r.mean()
x7_R=Werte_7_r.mean()
x8_R=Werte_8_r.mean()
x9_R=Werte_9_r.mean()
x10_R=Werte_10_r.mean()
x11_R=Werte_11_r.mean()
x12_R=Werte_12_r.mean()
x13_R=Werte_13_r.mean()
x14_R=Werte_14_r.mean()
x15_R=Werte_15_r.mean()
x16_R=Werte_16_r.mean()
x17_R=Werte_17_r.mean()



x1_M=Werte_1_m.mean()
x2_M=Werte_2_m.mean()
x3_M=Werte_3_m.mean()
x4_M=Werte_4_m.mean()
x5_M=Werte_5_m.mean()
x6_M=Werte_6_m.mean()
x7_M=Werte_7_m.mean()
x8_M=Werte_8_m.mean()
x9_M=Werte_9_m.mean()
x10_M=Werte_10_m.mean()
x11_M=Werte_11_m.mean()
x12_M=Werte_12_m.mean()
x13_M=Werte_13_m.mean()
x14_M=Werte_14_m.mean()
x15_M=Werte_15_m.mean()
x16_M=Werte_16_m.mean()
x17_M=Werte_17_m.mean()
X_Stufen=np.array([x1_M,x2_M,x3_M,x4_M,x5_M,x6_M,x7_M,x8_M,x9_M,x10_M,x11_M,x12_M,x13_M,x14_M,x15_M,x16_M,x17_M])

#%% Berechnung der totalen Mittelwerte


X1R=(x1_R+x17_R)/2
X2R=(x2_R+x16_R)/2
X3R=(x3_R+x15_R)/2
X4R=(x4_R+x14_R)/2
X5R=(x5_R+x13_R)/2
X6R=(x6_R+x12_R)/2
X7R=(x7_R+x11_R)/2
X8R=(x8_R+x10_R)/2
X9R=x9_R


XR=np.array([X1R,X2R,X3R,X4R,X5R,X6R,X7R,X8R,X9R])


X1M=(x1_M+x17_M)/2
X2M=(x2_M+x16_M)/2
X3M=(x3_M+x15_M)/2
X4M=(x4_M+x14_M)/2
X5M=(x5_M+x13_M)/2
X6M=(x6_M+x12_M)/2
X7M=(x7_M+x11_M)/2
X8M=(x8_M+x8_M)/2
X9M=x9_M


XM=np.array([X1M,X2M,X3M,X4M,X5M,X6M,X7M,X8M,X9M])

#%% Berechnung der von A

Laststufen=np.array([10,20,30,40,50,60,70,80,90])
print(XM-XR)
A = ((XM-XR ) / Laststufen) * 100


U=0.09

k=2


text=['-10%','-7.5%','-5%','-2.5%','0%','2.5%','5%','7.5%','10%']

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(abs(Laststufen),A,color='green')
plt.plot(abs(Laststufen),A-k*(U),color='blue',linestyle='--')
plt.plot(abs(Laststufen),A+k*(U),color='blue',linestyle='--')
plt.hlines(0,abs(Laststufen)[0],abs(Laststufen)[-1],colors='black')
plt.hlines(10,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.hlines(-10,abs(Laststufen)[0],abs(Laststufen)[-1],colors='red')
plt.yticks(np.arange(-10,11,2.5), text)
plt.grid()
plt.show()

