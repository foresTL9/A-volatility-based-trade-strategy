# coding: utf-8

import math
import xlrd
import xlwt
import numpy as np
import time
import matplotlib.pyplot as plt
from math import *

vics = 2850

mfuture = xlrd.open_workbook(u'/Zoom_2.xls').sheets()[0]
moption = xlrd.open_workbook(u'/Moption price_'+str(vics)+'.xlsx').sheets()[0]
dayvdata = xlrd.open_workbook(u'/dayvolmonths.xlsx').sheets()[0]
n1 = 4
n2 = 6
taryield = 0.01



def my_std(inplist, days):
    return [0 for i in range(days)]+[np.std(inplist[i:i+days], ddof = 1) for i in range(len(inplist)-days+1)]



def my_norm(inp):
    return np.exp(-inp*inp/2)/math.sqrt(2*3.141592653)

def phi(x):
    #'Cumulative distribution function for the standard normal distribution'
    return (1.0 + erf(x / sqrt(2.0))) / 2.0




def calc_vega(price_fut, price_st, vol, tau):
    d1 = (np.log(price_fut/vics)+(vol*vol/2)*tau)/(vol*np.sqrt(tau))
    vega = np.sqrt(tau)*price_fut*my_norm(d1)
    d1p = (np.log(price_st/vics)+(vol*vol/2)*tau)/(vol*np.sqrt(tau))
    delta = [phi(i) for i in d1p]
    return delta, vega



price_op = np.array(moption.col_values(1)[1:])
price_fut = np.array(mfuture.col_values(1)[1:])
t_date = np.array(mfuture.col_values(0)[1:])

price_st = np.array(dayvdata.col_values(1)[1:])
price_low = np.array(dayvdata.col_values(2)[1:])
tau = np.array(dayvdata.col_values(4)[1:])/60
day_vol = np.log(price_st/price_low)*np.log(price_st/price_low)/(4*np.log(2))




vol30 = np.array(mfuture.col_values(5)[1:])
delta, vega = calc_vega(price_fut,price_st,vol30,tau)
sigma = vega*day_vol




pos = 0
open_pos = 0
cyield = 0.00
tyield = [0]

sigv1 =[99999 for i in range(0,n2)] + [np.sum(sigma[i-n1:i])/np.sum(vega[i-n1:i]) for i in range(n2,len(sigma))]
sigv2=[99999 for i in range(0,n2)] + [np.sum(sigma[i-n2:i])/np.sum(vega[i-n1:i]) for i in range(n2,len(sigma))]

for i in range(n2, len(sigma)):
    if (cyield > taryield) or (sigv1[i]>sigv1[i-1]):
        cyield = 0
        pos = 0
    if pos == 0 and sigv1[i]<sigv2[i]:
        open_pos = 1
    if pos == 0 and open_pos == 1:
        open_pos = 0
        pos = 1
        continue
    if pos == 1:
        cyield += math.log((-price_op[i]+delta[i]*price_fut[i])/(-price_op[i-1]+delta[i]*price_fut[i-1]))
        tyield.append(tyield[len(tyield)-1]+(-price_op[i]+delta[i]*price_fut[i])-(-price_op[i-1]+delta[i]*price_fut[i-1]))
    else:
        tyield.append(tyield[len(tyield)-1])

    


output = xlwt.Workbook(encoding = 'ascii')
sheet1 = output.add_sheet('Sheet1',cell_overwrite_ok= True)
for i in range(0, len(tyield)):
    sheet1.write(i+6,1,tyield[i])
for i in range(0, len(vega)):
    sheet1.write(i+1,2,vega[i])
    sheet1.write(i+1,3,delta[i])
    sheet1.write(i+1,4,price_fut[i])
    sheet1.write(i+1,5,price_op[i])
    sheet1.write(i+1,6,day_vol[i])
    sheet1.write(i+1,7,sigv1[i])
    sheet1.write(i+1,8,sigv2[i])
output.save('/fm_c'+str(vics)+'vvol.xls')

