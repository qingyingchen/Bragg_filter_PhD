from xlrd import open_workbook#import data from Excel to python
import numpy as np
import matplotlib.pyplot as plt
#from scipy.interpolate import make_interp_spline
from scipy.signal import savgol_filter

x_data=[]
y_data=[]
y_data1=[]
y_data2=[]
y_data3=[]
y_data4=[]
y_data5=[]
y_data6=[]
x_volte=[]
temp=[]
wb = open_workbook\
     (r'C:/Users/qic656/Desktop/raw data/process data/1030-1120 smooth 4 m.xlsx')
 
for s in wb.sheets():
    for row in range(1250,2100,1):
        values = []
        for col in range(s.ncols):
            values.append(s.cell(row,col).value)
        x_data.append(values[0])
        y_data.append(values[4])
        y_data1.append(values[11])
        y_data2.append(values[18])
        y_data3.append(values[25])
        y_data4.append(values[32])
        y_data5.append(values[39])
        y_data6.append(values[49])

y_smooth = savgol_filter(y_data,53,3)#53 is the window_length, must be positive and odd; smaller, closer to the original shape;
#3 is the k, polyorder,should be smaller than window_length. smaller, smoothier.
y_1smooth = savgol_filter(y_data1,53,3)
y_2smooth = savgol_filter(y_data2,53,3)
y_3smooth = savgol_filter(y_data3,53,3)
y_4smooth = savgol_filter(y_data4,53,3)
y_5smooth = savgol_filter(y_data5,53,3)
y_6smooth = savgol_filter(y_data6,53,3)

figsize = 12,9
figure, ax = plt.subplots(figsize=figsize)
A,=plt.plot(x_data, y_smooth, color='dodgerblue',label=u"100 smooth",linewidth=3)
B,=plt.plot(x_data, y_1smooth, color='red',label=u"500 smooth",linewidth=3)
C,=plt.plot(x_data, y_2smooth, color='gray',label=u"1000 smooth",linewidth=3)
D,=plt.plot(x_data, y_3smooth, color='orange',label=u"2000 smooth",linewidth=3)
E,=plt.plot(x_data, y_4smooth, color='blue',label=u"2000S smooth",linewidth=3)
F,=plt.plot(x_data, y_5smooth, color='green',label=u"2000C smooth",linewidth=3)
G,=plt.plot(x_data, y_6smooth, color='purple',label=u"Dark normalized",linewidth=3)



plt.xlim((1055,1072))
plt.ylim((-50,5))

plt.xticks(np.arange(1056, 1072, 4.0))
plt.yticks(np.arange(-40, 5, 10.0))

#plt.title(u"Transmission")
font1 = {'family' : 'Arial', 'weight' : 'normal', 'size':20}
legend = plt.legend(handles=[A,B,C,D,E,F,G],prop=font1)

plt.tick_params(labelsize=30)
labels = ax.get_xticklabels() + ax.get_yticklabels()
[label.set_fontname('Arial') for label in labels]

font2 = {'family' : 'Arial', 'weight' : 'normal', 'size':35}
plt.xlabel(u"Wavelength (nm)",font2)
plt.ylabel(u"Transmission power (dB)",font2)
#plt.grid(linestyle='-.')
plt.grid()
plt.savefig('Savgol-filter smooth 1055-1072 with dark')
#plt.savefig(r'C:/Users/qic656/Desktop/raw data/process data/1040-1120 smooth savgol modify.png',dpi=500)
plt.show()


