# -*- coding: utf-8 -*-
"""
Created on Thu Jun 11 09:33:07 2020

@author: HI
"""

import xlwings as xw
import math
from scipy.stats import norm


@xw.func
def GBlackScholesNGreeek(S,X,T,r,q,sigam,corp):
    data=[S,X,T,r,q,sigam,corp]
    d1=(math.log(data[0]/data[1])+(data[4]+data[5]**2/2)*data[2]/365)/((data[5])*math.sqrt(data[2]/365))
    d2=d1-data[5]*math.sqrt(data[2]/365)
    c=data[0]*math.exp((data[4]-data[3])*data[2]/365)*norm.cdf(d1)-data[1]*math.exp(-data[3]*data[2]/365)*norm.cdf(d2)
    p=data[1]*math.exp(-data[3]*data[2]/365)*norm.cdf(-d2)-data[0]*math.exp((data[4]-data[3])*data[2]/365)*norm.cdf(-d1) 
    if data[6]=='c':
        return c
    else:
        return p
  





