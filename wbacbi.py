#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import numpy as np 
import pandas as pd
#from matplotlib import pyplot as plt 

class wbqual:
    def __init__(self,file):
        self.file=file
        self.nsh1=-5
        self.nsh2=-4
        self.nrowb_data=30
        self.nrowe_data=69
        self.ncol_data=7
        self.nrowb_param=25
        self.nrowe_param=28

    def pt(self,x,display=1,fexit=0):
        print("dir is:\n",dir(x))
        print("type is:\n ",type(x))
        if display==1: print("variable is:\n",x)
        if fexit!=0: exit(0)
        return x

    def makeratio(self,target='tmp.xlsx'):
        df1=pd.read_excel(self.file,sheet_name=self.nsh1)
        df2=pd.read_excel(self.file,sheet_name=self.nsh2)
        df1data=df1.iloc[self.nrowb_data:self.nrowe_data+1,self.ncol_data-1:]
        df2data=df2.iloc[self.nrowb_data:self.nrowe_data+1,self.ncol_data-1:]
        dfp=df1.iloc[self.nrowb_param-2:self.nrowe_param-1,self.ncol_data-1:]
        dfd=df2data
        for x in range(df1data.shape[0]):
            for y in range(df1data.shape[1]):
                d1=df1data.iloc[x,y]
                d2=df2data.iloc[x,y]
                if d1!=0 : dfd.iloc[x,y]=d2/d1
#                if d1!=0 : dfd.iloc[x,y]=d2
        dfdmax=dfd.max().apply(lambda x : '%.2f%%'%(x*100) if type(x)==float else '' ).to_frame().T
        dfdmin=dfd.min().apply(lambda x : '%.2f%%'%(x*100) if type(x)==float else '' ).to_frame().T
        dfdmean=dfd.mean().apply(lambda x :'%.2f%%'%(x*100)).to_frame().T
        dfd4sigmaper=(dfd.std()/dfd.mean()*4).apply(lambda x :'%.2f%%'%(x*100)).to_frame().T
        dfa=pd.concat([dfd,dfp,dfdmean,dfd4sigmaper,dfdmin,dfdmax])
        dfa.to_excel(target)

def main(file=sys.argv[1],target='tmp.xlsx'):
    q=wbqual(file)
    if len(sys.argv)>2: target=sys.argv[2]
    q.makeratio(target)
if __name__ == '__main__' : main()   
#if __name__ == '__main__' : main(sys.argv[1],sys.argv[2])   
