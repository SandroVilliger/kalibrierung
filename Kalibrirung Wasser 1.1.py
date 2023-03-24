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

file = r'C:\Users\Sandro Villiger\switchdrive\Arbeit IBI\Arbeit Uwe\037 Kallibrirung Killian\Kallibrirung der Wassermenge\2022-12-22-10_50_42.475_Kalibration_Wasser_1.1.csv'

#%%
# Import der Daten

data = pd.read_csv(file,sep=';')

Referenz=np.array(data['MW'])

Messung=np.array(data['Qw 11'])

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

#Max_1_1_1_r=Referenz[750:1396]
Max_1_1_2_r=Referenz[5292:5800]
#Max_1_R=np.concatenate((Max_1_1_1_r,Max_1_1_2_r),axis=0)
Max_1_R=Max_1_1_2_r

Max_0875_1_1_r=Referenz[4741:5126]
Max_0875_1_2_r=Referenz[6050:6307]
Max_0875_R=np.concatenate((Max_0875_1_1_r,Max_0875_1_2_r),axis=0)

Max_075_1_1_r=Referenz[4300:4670]
Max_075_1_2_r=Referenz[6405:6853]
Max_075_R=np.concatenate((Max_075_1_1_r,Max_075_1_2_r),axis=0)

Max_0625_1_1_r=Referenz[3673:4126]
Max_0625_1_2_r=Referenz[7000:7398]
Max_0625_R=np.concatenate((Max_0625_1_1_r,Max_0625_1_2_r),axis=0)

Max_05_1_1_r=Referenz[3200:3580]
Max_05_1_2_r=Referenz[7549:7962]
Max_05_R=np.concatenate((Max_05_1_1_r,Max_05_1_2_r),axis=0)

Max_0375_1_1_r=Referenz[2606:3034]
Max_0375_1_2_r=Referenz[8098:8490]
Max_0375_R=np.concatenate((Max_0375_1_1_r,Max_0375_1_2_r),axis=0)

Max_025_1_1_r=Referenz[2254:2488]
Max_025_1_2_r=Referenz[8561:9035]
Max_025_R=np.concatenate((Max_025_1_1_r,Max_025_1_2_r),axis=0)


Max_0125_1_1_r=Referenz[1750:2050]
Max_0125_1_2_r=Referenz[9141:9581]
Max_0125_R=np.concatenate((Max_0125_1_1_r,Max_0125_1_2_r),axis=0)

#Max_1_1_1_m=Messung[750:1396]
Max_1_1_2_m=Messung[5292:5800]
#Max_1_M=np.concatenate((Max_1_1_1_m,Max_1_1_2_m),axis=0)
Max_1_M=Max_1_1_2_m

Max_0875_1_1_m=Messung[4741:5126]
Max_0875_1_2_m=Messung[6050:6307]
Max_0875_M=np.concatenate((Max_0875_1_1_m,Max_0875_1_2_m),axis=0)

Max_075_1_1_m=Messung[4300:4670]
Max_075_1_2_m=Messung[6405:6853]
Max_075_M=np.concatenate((Max_075_1_1_m,Max_075_1_2_m),axis=0)

Max_0625_1_1_m=Messung[3673:4126]
Max_0625_1_2_m=Messung[7000:7398]
Max_0625_M=np.concatenate((Max_0625_1_1_m,Max_0625_1_2_m),axis=0)

Max_05_1_1_m=Messung[3200:3580]
Max_05_1_2_m=Messung[7549:7962]
Max_05_M=np.concatenate((Max_05_1_1_m,Max_05_1_2_m),axis=0)

Max_0375_1_1_m=Messung[2606:3034]
Max_0375_1_2_m=Messung[8098:8490]
Max_0375_M=np.concatenate((Max_0375_1_1_m,Max_0375_1_2_m),axis=0)

Max_025_1_1_m=Messung[2254:2488]
Max_025_1_2_m=Messung[8561:9035]
Max_025_M=np.concatenate((Max_025_1_1_m,Max_025_1_2_m),axis=0)

Max_0125_1_1_m=Messung[1750:2050]
Max_0125_1_2_m=Messung[9141:9581]
Max_0125_M=np.concatenate((Max_0125_1_1_m,Max_0125_1_2_m),axis=0)

#Punkte_1=np.concatenate((np.arange(750,1396),np.arange(5292,5800)),axis=0)
Punkte_1=np.arange(5292,5800)
Punkte_875=np.concatenate((np.arange(4741,5126),np.arange(6050,6307)),axis=0)
Punkte_75=np.concatenate((np.arange(4300,4670),np.arange(6405,6853)),axis=0)
Punkte_625=np.concatenate((np.arange(3673,4126),np.arange(7000,7398)),axis=0)
Punkte_05=np.concatenate((np.arange(3200,3580),np.arange(7549,7962)),axis=0)
Punkte_375=np.concatenate((np.arange(2606,3034),np.arange(8098,8490)),axis=0)
Punkte_25=np.concatenate((np.arange(2254,2488),np.arange(8561,9035)),axis=0)
Punkte_125=np.concatenate((np.arange(1750,2050),np.arange(9141,9581)),axis=0)




#%%


plt.figure(figsize=(8, 5), dpi=300)
plt.plot(Punkte,Referenz,color='black',label='Referenz')
plt.plot(Punkte,Referenz,color='blue',label='Messung')

#plt.plot(np.arange(750,1396),Max_1_1_1_r,color='red')
plt.plot(np.arange(5292,5800),Max_1_1_2_r,color='red',label='Auswertung')

plt.plot(np.arange(4741,5126),Max_0875_1_1_r,color='red')
plt.plot(np.arange(6050,6307),Max_0875_1_2_r,color='red')

plt.plot(np.arange(4300,4670),Max_075_1_1_r,color='red')
plt.plot(np.arange(6405,6853),Max_075_1_2_r,color='red')

plt.plot(np.arange(3673,4126),Max_0625_1_1_r,color='red')
plt.plot(np.arange(7000,7398),Max_0625_1_2_r,color='red')

plt.plot(np.arange(3200,3580),Max_05_1_1_r,color='red')
plt.plot(np.arange(7549,7962),Max_05_1_2_r,color='red')

plt.plot(np.arange(2606,3034),Max_0375_1_1_r,color='red')
plt.plot(np.arange(8098,8490),Max_0375_1_2_r,color='red')

plt.plot(np.arange(2254,2488),Max_025_1_1_r,color='red')
plt.plot(np.arange(8561,9035),Max_025_1_2_r,color='red')

plt.plot(np.arange(1750,2050),Max_0125_1_1_r,color='red')
plt.plot(np.arange(9141,9581),Max_0125_1_2_r,color='red')

plt.grid()
plt.xticks(np.arange(0,10001,1000))
plt.legend()
plt.show()


#%% Berechnung der Mittelwerte


x1_R=Max_0125_1_1_r.mean()
x2_R=Max_025_1_1_r.mean()
x3_R=Max_0375_1_1_r.mean()
x4_R=Max_05_1_1_r.mean()
x5_R=Max_0625_1_1_r.mean()
x6_R=Max_075_1_1_r.mean()
x7_R=Max_0875_1_1_r.mean()
x8_R=Max_1_1_2_r.mean()
x9_R=Max_0875_1_2_r.mean()
x10_R=Max_075_1_2_r.mean()
x11_R=Max_0625_1_2_r.mean()
x12_R=Max_05_1_2_r.mean()
x13_R=Max_0375_1_2_r.mean()
x14_R=Max_025_1_2_r.mean()
x15_R=Max_0125_1_2_r.mean()


x1_M=Max_0125_1_1_m.mean()
x2_M=Max_025_1_1_m.mean()
x3_M=Max_0375_1_1_m.mean()
x4_M=Max_05_1_1_m.mean()
x5_M=Max_0625_1_1_m.mean()
x6_M=Max_075_1_1_m.mean()
x7_M=Max_0875_1_1_m.mean()
x8_M=Max_1_1_2_m.mean()
x9_M=Max_0875_1_2_m.mean()
x10_M=Max_075_1_2_m.mean()
x11_M=Max_0625_1_2_m.mean()
x12_M=Max_05_1_2_m.mean()
x13_M=Max_0375_1_2_m.mean()
x14_M=Max_025_1_2_m.mean()
x15_M=Max_0125_1_2_m.mean()

X_Stufen=np.array([x1_M,x2_M,x3_M,x4_M,x5_M,x6_M,x7_M,x8_M,x9_M,x10_M,x11_M,x12_M,x13_M,x14_M,x15_M])

#%% Berechnung der totalen Mittelwerte


X1R=(x1_R+x15_R)/2
X2R=(x2_R+x14_R)/2
X3R=(x3_R+x13_R)/2
X4R=(x4_R+x12_R)/2
X5R=(x5_R+x11_R)/2
X6R=(x6_R+x10_R)/2
X7R=(x7_R+x9_R)/2
X8R=x8_R
XR=np.array([X1R,X2R,X3R,X4R,X5R,X6R,X7R,X8R])


X1M=(x1_M+x15_M)/2
X2M=(x2_M+x14_M)/2
X3M=(x3_M+x13_M)/2
X4M=(x4_M+x12_M)/2
X5M=(x5_M+x11_M)/2
X6M=(x6_M+x10_M)/2
X7M=(x7_M+x9_M)/2
X8M=x8_M

XM=np.array([X1M,X2M,X3M,X4M,X5M,X6M,X7M,X8M])

#%% Berechnung der von A

Laststufen=np.array([1,5,10,15,20,25,30,35])
print(XM-XR)
A = ((XM-XR ) / Laststufen) * 100

U=0.09

k=2

text=['-10%','-7.5%','-5%','-2.5%','0%','2.5%','5%','7.5%','10%']

plt.figure(figsize=(8, 5), dpi=300)
plt.plot(abs(Laststufen[1:]),A[1:],color='green')
plt.plot(abs(Laststufen[1:]),A[1:]-k*(U),color='blue',linestyle='--')
plt.plot(abs(Laststufen[1:]),A[1:]+k*(U),color='blue',linestyle='--')
plt.hlines(0,abs(Laststufen)[1],abs(Laststufen)[-1],colors='black')
plt.hlines(10,abs(Laststufen)[1],abs(Laststufen)[-1],colors='red')
plt.hlines(-10,abs(Laststufen)[1],abs(Laststufen)[-1],colors='red')
plt.yticks(np.arange(-10,11,2.5), text)
plt.grid()
plt.show()